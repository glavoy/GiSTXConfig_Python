from __future__ import annotations

import re
from pathlib import Path

from models import CalculationPart, CalculationType, Question, ResponseSourceType

class XmlGenerator:
    def __init__(self) -> None:
        self.logstring: list[str] = []

    def write_xml(self, worksheet_name: str, question_list: list[Question], xml_path: Path) -> Path:
        if worksheet_name.endswith("_dd"):
            xml_name = worksheet_name[:-3]
        else:
            xml_name = worksheet_name[:-4]

        out_file = xml_path / f"{xml_name}.xml"
        with out_file.open("w", encoding="utf-8", newline="\n") as f:
            f.write("<?xml version = '1.0' encoding = 'utf-8'?>\n")
            f.write("<survey>\n\n")

            for q in question_list:
                f.write(
                    f"\t<question type = '{q.questionType}' fieldname = '{q.fieldName}' fieldtype = '{q.fieldType}'>\n"
                )

                if q.questionType != "automatic":
                    f.write(f"\t\t<text>{q.questionText}</text>\n")

                if q.questionType == "automatic" and q.calculationType != CalculationType.NONE:
                    self._generate_calculation_xml(f, q)

                if q.maxCharacters != "-9":
                    f.write(f"\t\t<maxCharacters>{q.maxCharacters}</maxCharacters>\n")

                if q.mask:
                    f.write(f"\t\t<mask value=\"{q.mask}\" />\n")

                if q.uniqueCheckMessage:
                    f.write("\t\t<unique_check>\n")
                    f.write(f"\t\t\t<message>{q.uniqueCheckMessage}</message>\n")
                    f.write("\t\t</unique_check>\n")

                if q.questionType != "date" and q.lowerRange != "-9":
                    f.write("\t\t<numeric_check>\n")
                    f.write(
                        f"\t\t\t<values minvalue ='{q.lowerRange}' maxvalue='{q.upperRange}' other_values = '{q.lowerRange}' "
                        f"message = 'Number must be between {q.lowerRange} and {q.upperRange}!'></values>\n"
                    )
                    f.write("\t\t</numeric_check>\n")

                if q.questionType == "date":
                    f.write("\t\t<date_range>\n")
                    f.write(f"\t\t\t<min_date>{q.lowerRange}</min_date>\n")
                    f.write(f"\t\t\t<max_date>{q.upperRange}</max_date>\n")
                    f.write("\t\t</date_range>\n")

                for logic_check in q.logicChecks:
                    f.write("\t\t<logic_check>\n")
                    f.write(self._generate_logic_check(logic_check) + "\n")
                    f.write("\t\t</logic_check>\n")

                if q.questionType in {"radio", "checkbox", "combobox"}:
                    attrs = ""
                    if q.responseSourceType == ResponseSourceType.CSV:
                        attrs += f" source='csv' file='{q.responseSourceFile}'"
                    elif q.responseSourceType == ResponseSourceType.DATABASE:
                        attrs += f" source='database' table='{q.responseSourceTable}'"
                    f.write(f"\t\t<responses{attrs}>\n")

                    for flt in q.responseFilters:
                        f.write(
                            f"\t\t\t<filter column='{flt.column}' operator='{flt.operator}' value='{flt.value}'/>\n"
                        )
                    if q.responseDisplayColumn:
                        f.write(f"\t\t\t<display column='{q.responseDisplayColumn}'/>\n")
                    if q.responseValueColumn:
                        f.write(f"\t\t\t<value column='{q.responseValueColumn}'/>\n")
                    if q.responseDistinct is not None:
                        f.write(f"\t\t\t<distinct>{str(q.responseDistinct).lower()}</distinct>\n")
                    if q.responseEmptyMessage:
                        f.write(f"\t\t\t<empty_message>{q.responseEmptyMessage}</empty_message>\n")
                    if q.responseDontKnowValue:
                        label_attr = f" label='{q.responseDontKnowLabel}'" if q.responseDontKnowLabel else ""
                        f.write(f"\t\t\t<dont_know value='{q.responseDontKnowValue}'{label_attr}/>\n")
                    if q.responseNotInListValue:
                        label_attr = f" label='{q.responseNotInListLabel}'" if q.responseNotInListLabel else ""
                        f.write(f"\t\t\t<not_in_list value='{q.responseNotInListValue}'{label_attr}/>\n")

                    if q.responseSourceType == ResponseSourceType.STATIC:
                        responses = [r for r in re.split(r"\r\n|\n|\r", q.responses) if r]
                        if len(responses) == 0:
                            f.write("\t\t\t<response></response>\n")
                        else:
                            for response in responses:
                                index = response.find(":")
                                value = response[:index]
                                label = response[index + 1 :].strip()
                                f.write(f"\t\t\t<response value = '{value}'>{label}</response>\n")
                    f.write("\t\t</responses>\n")

                if q.skip:
                    skips = [s for s in re.split(r"\r\n|\n|\r", q.skip) if s]
                    pre = [s for s in skips if s.startswith("preskip:")]
                    post = [s for s in skips if s.startswith("postskip:")]
                    if pre:
                        f.write("\t\t<preskip>\n")
                        for s in pre:
                            f.write(self._generate_skip(s, "preSkip") + "\n")
                        f.write("\t\t</preskip>\n")
                    if post:
                        f.write("\t\t<postskip>\n")
                        for s in post:
                            f.write(self._generate_skip(s, "postSkip") + "\n")
                        f.write("\t\t</postskip>\n")

                if q.dontKnow in {"TRUE", "True"}:
                    f.write("\t\t<dont_know>-7</dont_know>\n")
                if q.refuse in {"TRUE", "True"}:
                    f.write("\t\t<refuse>-8</refuse>\n")
                if q.na in {"TRUE", "True"}:
                    f.write("\t\t<na>-6</na>\n")

                f.write("\t</question>\n\n")

            f.write("\t<question type = 'information' fieldname = 'end_of_questions' fieldtype = 'n/a'>\n")
            f.write("\t\t<text>Press the 'Finish' button to save the data.</text >\n")
            f.write("\t</question>\n\n")
            f.write("</survey>\n")

        return out_file

    def _generate_skip(self, skip: str, skip_type: str) -> str:
        len_skip = 13 if skip_type == "postSkip" else 12
        space_indices = [i for i, ch in enumerate(skip) if ch == " "]
        fieldname_to_check = skip[len_skip : space_indices[2]]

        if len(space_indices) == 9:
            condition = "does not contain"
            value = skip[space_indices[5] + 1 : space_indices[6] - 1]
        elif "contains" in skip:
            condition = "contains"
            value = skip[space_indices[3] + 1 : space_indices[4] - 1]
        else:
            condition = skip[space_indices[2] + 1 : space_indices[3]]
            condition = condition.replace("<", "&lt;").replace(">", "&gt;")
            value = skip[space_indices[3] + 1 : space_indices[4] - 1]

        fieldname_to_skip_to = skip[space_indices[-1] + 1 :]
        return (
            f"\t\t\t<skip fieldname='{fieldname_to_check}' condition = '{condition}' response='{value}' "
            f"response_type='fixed' skiptofieldname ='{fieldname_to_skip_to}'></skip>"
        )

    def _generate_logic_check(self, logic_check: str) -> str:
        expression, message = [p.strip() for p in logic_check.split(";", 1)]
        expression = expression.replace("!=", "&lt;&gt;")
        expression = expression.replace("<>", "&lt;&gt;")
        expression = expression.replace("<=", "&lt;=")
        expression = expression.replace(">=", "&gt;=")
        expression = re.sub(r"(?<!&lt;)(?<!&gt;)<(?!=)", "&lt;", expression)
        expression = re.sub(r"(?<!&lt;=)(?<!&gt;=)>(?!=)", "&gt;", expression)

        if " or " in expression:
            parts = expression.split(" or ")
            lines = []
            for i, part in enumerate(parts):
                suffix = " or" if i < len(parts) - 1 else ""
                lines.append(f"\t\t\t{part.strip()}{suffix}")
            lines.append(";")
            lines.append(f"\t\t\t{message}")
            return "\n".join(lines)

        return f"\t\t\t{expression}; {message}"

    def _generate_calculation_xml(self, f, q: Question) -> None:
        if q.calculationType == CalculationType.QUERY:
            f.write("\t\t<calculation type='query'>\n")
            f.write(f"\t\t\t<sql>{q.calculationQuerySql}</sql>\n")
            for param in q.calculationQueryParameters:
                f.write(f"\t\t\t<parameter name='{param.name}' field='{param.fieldName}' />\n")
            f.write("\t\t</calculation>\n")
        elif q.calculationType == CalculationType.CASE:
            f.write("\t\t<calculation type='case'>\n")
            for cond in q.calculationCaseConditions:
                op = self._convert_operator_to_xml(cond.operator)
                f.write(f"\t\t\t<when field='{cond.field}' operator='{op}' value='{cond.value}'>\n")
                if cond.result:
                    self._generate_calculation_part(f, cond.result, 4)
                f.write("\t\t\t</when>\n")
            if q.calculationCaseElse:
                f.write("\t\t\t<else>\n")
                self._generate_calculation_part(f, q.calculationCaseElse, 4)
                f.write("\t\t\t</else>\n")
            f.write("\t\t</calculation>\n")
        elif q.calculationType == CalculationType.CONSTANT:
            f.write(f"\t\t<calculation type='constant' value='{q.calculationConstantValue}' />\n")
        elif q.calculationType == CalculationType.LOOKUP:
            f.write(f"\t\t<calculation type='lookup' field='{q.calculationLookupField}' />\n")
        elif q.calculationType == CalculationType.MATH:
            f.write(f"\t\t<calculation type='math' operator='{q.calculationMathOperator}'>\n")
            for part in q.calculationMathParts:
                self._generate_calculation_part(f, part, 3)
            f.write("\t\t</calculation>\n")
        elif q.calculationType == CalculationType.CONCAT:
            separator_attr = f" separator='{q.calculationConcatSeparator}'" if q.calculationConcatSeparator else ""
            f.write(f"\t\t<calculation type='concat'{separator_attr}>\n")
            for part in q.calculationConcatParts:
                self._generate_calculation_part(f, part, 3)
            f.write("\t\t</calculation>\n")
        elif q.calculationType == CalculationType.AGE_FROM_DATE:
            f.write(
                f"\t\t<calculation type='age_from_date' field='{q.calculationLookupField}' value='{q.calculationConstantValue}'/>\n"
            )
        elif q.calculationType == CalculationType.AGE_AT_DATE:
            separator_attr = f" separator='{q.calculationConcatSeparator}'" if q.calculationConcatSeparator else ""
            f.write(
                f"\t\t<calculation type='age_at_date' field='{q.calculationLookupField}' value='{q.calculationConstantValue}'{separator_attr}/>\n"
            )
        elif q.calculationType == CalculationType.DATE_OFFSET:
            f.write(
                f"\t\t<calculation type='date_offset' field='{q.calculationLookupField}' value='{q.calculationConstantValue}' />\n"
            )
        elif q.calculationType == CalculationType.DATE_DIFF:
            f.write(
                f"\t\t<calculation type='date_diff' field='{q.calculationLookupField}' value='{q.calculationConstantValue}' unit='{q.calculationUnit}' />\n"
            )

    def _generate_calculation_part(self, f, part: CalculationPart, indent_level: int) -> None:
        indent = "\t" * indent_level
        if part.type == CalculationType.CONSTANT:
            f.write(f"{indent}<result type='constant' value='{part.constantValue}' />\n")
        elif part.type == CalculationType.LOOKUP:
            f.write(f"{indent}<part type='lookup' field='{part.lookupField}' />\n")
        elif part.type == CalculationType.QUERY:
            f.write(f"{indent}<part type='query'>\n")
            f.write(f"{indent}\t<sql>{part.querySql}</sql>\n")
            for param in part.queryParameters:
                f.write(f"{indent}\t<parameter name='{param.name}' field='{param.fieldName}' />\n")
            f.write(f"{indent}</part>\n")
        elif part.type == CalculationType.MATH:
            f.write(f"{indent}<part type='math' operator='{part.mathOperator}'>\n")
            for nested in part.parts:
                self._generate_calculation_part(f, nested, indent_level + 1)
            f.write(f"{indent}</part>\n")
        elif part.type == CalculationType.CONCAT:
            separator_attr = f" separator='{part.concatSeparator}'" if part.concatSeparator else ""
            f.write(f"{indent}<part type='concat'{separator_attr}>\n")
            for nested in part.parts:
                self._generate_calculation_part(f, nested, indent_level + 1)
            f.write(f"{indent}</part>\n")

    @staticmethod
    def _convert_operator_to_xml(op: str) -> str:
        op = op.strip()
        return {
            "=": "=",
            "!=": "!=",
            "<>": "&lt;&gt;",
            ">": "&gt;",
            "<": "&lt;",
            ">=": "&gt;=",
            "<=": "&lt;=",
        }.get(op, "=")
