import re
import unicodedata

class SmartTfnEngine:
    @staticmethod
    def extract_digits(input_str: str) -> str:
        """Extract only digit characters from string."""
        if not input_str:
            return ""
        return re.sub(r'\D', '', input_str)

    @classmethod
    def format_to_competitor_pattern(cls, competitor_str: str, target_input_str: str) -> str:
        """
        Maps target digits/country code onto competitor's exact formatting template.
        Preserves non-digit symbols (+, *, (, ), -, ., spaces).
        """
        if not competitor_str or not target_input_str:
            return competitor_str

        target_digits = cls.extract_digits(target_input_str)
        if not target_digits:
            return competitor_str

        result = []
        target_idx = 0

        for char in competitor_str:
            normalized = unicodedata.normalize('NFKC', char)
            if normalized.isdigit():
                if target_idx < len(target_digits):
                    result.append(target_digits[target_idx])
                    target_idx += 1
            else:
                result.append(char) # Preserve symbol

        # Append remaining digits if target was longer than competitor slots
        while target_idx < len(target_digits):
            result.append(target_digits[target_idx])
            target_idx += 1

        return "".join(result)

    @classmethod
    def replace_text_alternating(cls, text: str, target_lines: list) -> tuple:
        """
        Scans text for phone numbers and replaces them in round-robin alternating order.
        Returns (new_text, count_replaced).
        """
        if not target_lines:
            return text, 0

        tfn_regex = re.compile(
            r'(?:\+?[^\w\s\n]{0,4})?(?:\d[^\w\s\n]{0,4}){9,14}\d',
            re.UNICODE
        )

        matches = list(tfn_regex.finditer(text))
        if not matches:
            return text, 0

        result_pieces = []
        last_idx = 0
        occurrence_idx = 0

        for match in matches:
            start, end = match.span()
            result_pieces.append(text[last_idx:start])

            competitor_number = match.group(0)
            target_input = target_lines[occurrence_idx % len(target_lines)]
            occurrence_idx += 1

            formatted_number = cls.format_to_competitor_pattern(competitor_number, target_input)
            result_pieces.append(formatted_number)

            last_idx = end

        result_pieces.append(text[last_idx:])
        return "".join(result_pieces), len(matches)
