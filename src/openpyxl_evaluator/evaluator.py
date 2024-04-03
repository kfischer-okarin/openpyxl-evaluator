from dataclasses import dataclass
from datetime import datetime
from functools import cached_property
from typing import Any

from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.worksheet.worksheet import Worksheet


class Evaluator:
    def __init__(self, cell):
        self.cell = cell
        self._parsed_expressions = []

    @cached_property
    def value(self):
        try:
            if self._is_formula:
                return self._evaluate_formula()
        except NotImplementedError as e:
            raise EvaluationError(self.cell) from e

        return self.cell.value

    @property
    def _is_formula(self):
        return self.cell.data_type == "f"

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
        elif _is_infix_operator(next_token):
            self._parsed_expressions.append(self._consume_infix_operator(tokens))
        else:
            error_message = f"Token {next_token} not yet implemented"
            raise NotImplementedError(error_message)

    def _consume_function(self, tokens):
        next_token = tokens.pop(0)
        name = next_token.value[:-1]  # Remove the opening parenthesis
        operands = []
        while not _is_function_end(tokens[0]):
            self._consume_function_operand(tokens)
            operands.append(self._parsed_expressions.pop())

        tokens.pop(0)  # Remove the closing parenthesis

        return Function(name, operands)

    def _consume_function_operand(self, tokens):
        while True:
            self._consume_next_expression(tokens)

            if _is_function_end(tokens[0]):
                break

            if _is_argument_separator(tokens[0]):
                self._consume_argument_separator(tokens)
                break

    def _consume_argument_separator(self, tokens):
        while _is_argument_separator(tokens[0]):
            tokens.pop(0)

    def _consume_operand(self, tokens):
        next_token = tokens.pop(0)

        if next_token.subtype == Token.TEXT:
            return Value(next_token.value.strip('"'))

        if next_token.subtype == Token.NUMBER:
            try:
                return Value(int(next_token.value))
            except ValueError:
                return Value(float(next_token.value))

        if next_token.subtype == Token.RANGE:
            return Range(self.cell.parent, next_token.value)

        error_message = f"Operand {next_token} not yet implemented"
        raise NotImplementedError(error_message)

    def _consume_infix_operator(self, tokens):
        left = self._parsed_expressions.pop()
        operator = tokens.pop(0).value
        right = self._consume_operand(tokens)

        return InfixOperator(left, operator, right)


class EvaluationError(Exception):
    def __init__(self, cell):
        super().__init__(f"Error evaluating cell '{cell}'")


@dataclass(frozen=True)
class Value:
    value: Any

    def evaluate(self):
        return self.value


@dataclass(frozen=True)
class Range:
    worksheet: Worksheet
    reference: str

    def evaluate(self):
        target_range = self.worksheet[self.reference]
        if isinstance(target_range, tuple):
            return tuple(tuple(Evaluator(cell).value for cell in row) for row in target_range)

        return Evaluator(self.worksheet[self.reference]).value


@dataclass(frozen=True)
class InfixOperator:
    left: Any
    operator: str
    right: Any

    def evaluate(self):
        if self.operator == "+":
            return self.left.evaluate() + self.right.evaluate()

        if self.operator == "-":
            return self.left.evaluate() - self.right.evaluate()

        if self.operator == "*":
            return self.left.evaluate() * self.right.evaluate()

        if self.operator == "/":
            return self.left.evaluate() / self.right.evaluate()

        error_message = f"Unknown Infix Operator: {self.operator}"
        raise ValueError(error_message)


@dataclass(frozen=True)
class Function:
    name: str
    operands: list

    def evaluate(self):
        if self.name == "DATEVALUE":
            return datetime.strptime(self.operands[0].evaluate(), "%Y-%m-%d").date()  # noqa: DTZ007

        if self.name == "SUM":
            return sum(
                sum(cell or 0 for cell in row)  # Empty cells (None) are treated as 0
                for row in self.operands[0].evaluate()
            )

        if self.name == "IFERROR":
            try:
                return self.operands[0].evaluate()
            except ZeroDivisionError:
                return self.operands[1].evaluate()

        error_message = f"Function {self.name} not yet implemented"
        raise NotImplementedError(error_message)


def _is_function_start(token):
    return token.type == Token.FUNC and token.subtype == Token.OPEN


def _is_function_end(token):
    return token.type == Token.FUNC and token.subtype == Token.CLOSE


def _is_argument_separator(token):
    return token.type in {Token.SEP, Token.WSPACE} or token.value == ","


def _is_operand(token):
    return token.type == Token.OPERAND


def _is_infix_operator(token):
    return token.type == Token.OP_IN
