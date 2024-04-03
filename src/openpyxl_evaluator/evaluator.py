from dataclasses import dataclass
from datetime import datetime
from functools import cached_property

from openpyxl.formula.tokenizer import Tokenizer, Token


class Evaluator:
    def __init__(self, cell):
        self.cell = cell
        self._parsed_expressions = []

    @cached_property
    def value(self):
        try:
            if self._is_formula:
                return self._evaluate_formula()

            return self.cell.value
        except Exception as e:
            raise EvaluationError(f"Error evaluating cell '{self.cell.value}': {e}") from e

    @property
    def _is_formula(self):
        return self.cell.data_type == 'f'

    def _evaluate_formula(self):
        tokens = Tokenizer(self.cell.value).items

        while tokens:
            self._consume_next_expression(tokens)

        result = self._parsed_expressions.pop()

        return result.evaluate()

    def _consume_next_expression(self, tokens):
        next_token = tokens[0]

        if _is_function_start(next_token):
            self._parsed_expressions.append(self._consume_function(tokens))
        elif _is_operand(next_token):
            self._parsed_expressions.append(self._consume_operand(tokens))
        else:
            raise NotImplementedError(f"Token {next_token} not yet implemented")

    def _consume_function(self, tokens):
        next_token = tokens.pop(0)
        name = next_token.value[:-1] # Remove the opening parenthesis
        operands = []
        while not _is_function_end(tokens[0]):
            self._consume_function_operand(tokens)
            operands.append(self._parsed_expressions.pop())

        tokens.pop(0) # Remove the closing parenthesis

        return Function(name, operands)

    def _consume_function_operand(self, tokens):
        while True:
            self._consume_next_expression(tokens)

            if _is_function_end(tokens[0]):
                break

    def _consume_operand(self, tokens):
        next_token = tokens.pop(0)

        if next_token.subtype == Token.TEXT:
            return Value(next_token.value.strip('"'))

        raise NotImplementedError(f"Operand {next_token} not yet implemented")


class EvaluationError(Exception):
    pass


@dataclass
class Value:
    value: str

    def evaluate(self):
        return self.value


@dataclass
class Function:
    name: str
    operands: list

    def evaluate(self):
        if self.name == 'DATEVALUE':
            return datetime.strptime(self.operands[0].evaluate(), '%Y-%m-%d').date()

        raise NotImplementedError(f"Function {self.name} not yet implemented")


def _is_function_start(token):
    return token.type == Token.FUNC and token.subtype == Token.OPEN

def _is_function_end(token):
    return token.type == Token.FUNC and token.subtype == Token.CLOSE

def _is_argument_separator(token):
    return token.type == Token.SEP

def _is_operand(token):
    return token.type == Token.OPERAND
