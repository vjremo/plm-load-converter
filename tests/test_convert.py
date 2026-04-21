import pytest
from convert import build_lookup, parse_header, transform_cell

# Minimal lookup table used across all list-based tests
LOOKUP = {
    "Department": {"Dept 1": "xyzDept1", "Dept 2": "xyzDept2"},
    "Class":      {"Class 1": "xyzClass1", "Class 2": "xyzClass2"},
    "Age Range":  {"0 to 3 years": "xyz0to3", "3 to 6 years": "xyz4to6", "6 plus": "xyz6plus"},
}


# ── parse_header ─────────────────────────────────────────────────────────────

class TestParseHeader:
    def test_single_list(self):
        assert parse_header("Department-SingleList") == ("Department", "SingleList")

    def test_color_choice_maps_to_single_list(self):
        assert parse_header("Color-ColorChoice") == ("Color", "SingleList")

    def test_multi_list(self):
        assert parse_header("Age Range-MultiList") == ("Age Range", "MultiList")

    def test_composite_maps_to_multi_list(self):
        assert parse_header("Style-Composite") == ("Style", "MultiList")

    def test_float(self):
        assert parse_header("Price-Float") == ("Price", "Float")

    def test_boolean(self):
        assert parse_header("Active-Boolean") == ("Active", "Boolean")

    def test_integer(self):
        assert parse_header("Sort Order-Integer") == ("Sort Order", "Integer")

    def test_multi_entry(self):
        assert parse_header("Tags-MultiEntry") == ("Tags", "MultiEntry")

    def test_plain_no_suffix(self):
        assert parse_header("Object Type") == ("Object Type", None)

    def test_plain_with_hyphen_but_unknown_suffix(self):
        assert parse_header("Some-Unknown") == ("Some-Unknown", None)


# ── SingleList ────────────────────────────────────────────────────────────────

class TestSingleList:
    def test_resolves_display_to_internal(self):
        assert transform_cell("Dept 1", "SingleList", "Department", LOOKUP, 1, "col") == "xyzDept1"

    def test_resolves_second_entry(self):
        assert transform_cell("Dept 2", "SingleList", "Department", LOOKUP, 1, "col") == "xyzDept2"

    def test_unknown_value_raises(self):
        with pytest.raises(ValueError, match="not found in References"):
            transform_cell("Dept 99", "SingleList", "Department", LOOKUP, 1, "col")

    def test_unknown_category_raises(self):
        with pytest.raises(ValueError, match="no reference table"):
            transform_cell("X", "SingleList", "NoSuchCategory", LOOKUP, 1, "col")

    def test_color_choice_suffix_resolves(self):
        _, suffix = parse_header("Class-ColorChoice")
        assert transform_cell("Class 1", suffix, "Class", LOOKUP, 1, "col") == "xyzClass1"


# ── MultiList ─────────────────────────────────────────────────────────────────

class TestMultiList:
    def test_single_value(self):
        assert transform_cell("6 plus", "MultiList", "Age Range", LOOKUP, 1, "col") == "xyz6plus"

    def test_two_values(self):
        result = transform_cell("3 to 6 years,6 plus", "MultiList", "Age Range", LOOKUP, 1, "col")
        assert result == "xyz4to6|~*~|xyz6plus"

    def test_three_values(self):
        result = transform_cell("0 to 3 years,3 to 6 years,6 plus", "MultiList", "Age Range", LOOKUP, 1, "col")
        assert result == "xyz0to3|~*~|xyz4to6|~*~|xyz6plus"

    def test_whitespace_around_commas(self):
        result = transform_cell("3 to 6 years , 6 plus", "MultiList", "Age Range", LOOKUP, 1, "col")
        assert result == "xyz4to6|~*~|xyz6plus"

    def test_unknown_value_raises(self):
        with pytest.raises(ValueError, match="not found in References"):
            transform_cell("6 plus,Unknown", "MultiList", "Age Range", LOOKUP, 1, "col")

    def test_composite_suffix_resolves(self):
        _, suffix = parse_header("Age Range-Composite")
        result = transform_cell("3 to 6 years,6 plus", suffix, "Age Range", LOOKUP, 1, "col")
        assert result == "xyz4to6|~*~|xyz6plus"


# ── Float ─────────────────────────────────────────────────────────────────────

class TestFloat:
    @pytest.mark.parametrize("value", ["3.05", "0", "100", "-1.5", "0.001"])
    def test_valid_float(self, value):
        assert transform_cell(value, "Float", None, LOOKUP, 1, "col") == value

    @pytest.mark.parametrize("value", ["abc", "3.0.0", "", "1e2x"])
    def test_invalid_float_raises(self, value):
        with pytest.raises(ValueError, match="not a valid floating-point number"):
            transform_cell(value, "Float", None, LOOKUP, 1, "col")


# ── Boolean ───────────────────────────────────────────────────────────────────

class TestBoolean:
    def test_yes_becomes_true(self):
        assert transform_cell("Yes", "Boolean", None, LOOKUP, 1, "col") == "true"

    def test_no_becomes_false(self):
        assert transform_cell("No", "Boolean", None, LOOKUP, 1, "col") == "false"

    @pytest.mark.parametrize("value", ["yes", "no", "TRUE", "1", "maybe", ""])
    def test_invalid_boolean_raises(self, value):
        with pytest.raises(ValueError, match="Boolean field must be"):
            transform_cell(value, "Boolean", None, LOOKUP, 1, "col")


# ── Integer ───────────────────────────────────────────────────────────────────

class TestInteger:
    @pytest.mark.parametrize("value", ["0", "1", "42", "100", "-5"])
    def test_valid_integer(self, value):
        assert transform_cell(value, "Integer", None, LOOKUP, 1, "col") == value

    @pytest.mark.parametrize("value", ["3.5", "3.0", "abc", "", "1e2"])
    def test_invalid_integer_raises(self, value):
        with pytest.raises(ValueError, match="not a valid whole number"):
            transform_cell(value, "Integer", None, LOOKUP, 1, "col")


# ── MultiEntry ────────────────────────────────────────────────────────────────

class TestMultiEntry:
    def test_single_value(self):
        assert transform_cell("Solo", "MultiEntry", None, LOOKUP, 1, "col") == "Solo"

    def test_two_values(self):
        assert transform_cell("Test1, Test2", "MultiEntry", None, LOOKUP, 1, "col") == "Test1|~*~|Test2"

    def test_three_values(self):
        assert transform_cell("A,B,C", "MultiEntry", None, LOOKUP, 1, "col") == "A|~*~|B|~*~|C"

    def test_trims_whitespace(self):
        assert transform_cell("  Foo , Bar  ", "MultiEntry", None, LOOKUP, 1, "col") == "Foo|~*~|Bar"

    def test_no_reference_lookup(self):
        # values that don't exist in LOOKUP should still pass — no lookup performed
        assert transform_cell("Unknown1,Unknown2", "MultiEntry", None, LOOKUP, 1, "col") == "Unknown1|~*~|Unknown2"


# ── Plain (no suffix) ─────────────────────────────────────────────────────────

class TestPlain:
    def test_string_passthrough(self):
        assert transform_cell("Product", None, None, LOOKUP, 1, "col") == "Product"

    def test_none_becomes_empty_string(self):
        assert transform_cell(None, None, None, LOOKUP, 1, "col") == ""

    def test_numeric_value_as_string(self):
        assert transform_cell(42, None, None, LOOKUP, 1, "col") == "42"
