//! Literal forms and type tags — Lua 5.x manual §3.1–3.3.

lua_print! {
hex_literal_evaluates => { "print(0x10)\n", "16" },
hex_uppercase_prefix => { "print(0XFF)\n", "255" },
scientific_notation_integer => { "print(1e2)\n", "100" },
scientific_notation_fractional_mantissa => { "print(2.5e1)\n", "25" },
long_bracket_string_preserves_backslash => { "print([[\\n]])\n", "\\n" },
double_quoted_escape_newline => { "print(\"a\\nb\")\n", "a\nb" },
double_quoted_escape_tab => { "print(\"a\\tb\")\n", "a\tb" },
double_quoted_escape_backslash => { "print(\"\\\\\")\n", "\\" },
double_quoted_escape_double_quote => { "print(\"\\\"\")\n", "\"" },
single_quoted_string_literal => { "print('lua')\n", "lua" },
type_of_nil => { "print(type(nil))\n", "nil" },
type_of_boolean => { "print(type(false))\n", "boolean" },
type_of_number => { "print(type(0))\n", "number" },
type_of_string => { "print(type(\"\"))\n", "string" },
type_of_table => { "print(type({}))\n", "table" },
type_of_function => { "print(type(function() end))\n", "function" },
empty_table_is_truthy => { "print(not not {})\n", "true" },
empty_string_is_truthy => { "print(not not \"\")\n", "true" },
false_is_falsy => { "print(not false)\n", "true" },
nil_is_falsy => { "print(not not nil)\n", "false" },
binary_literal_evaluates => { "print(0b1010)\n", "10" },
binary_literal_with_uppercase_prefix => { "print(0B11)\n", "3" },
decimal_point_without_leading_digit => { "print(.5 * 2)\n", "1" },
decimal_point_without_trailing_digit => { "print(1. * 2)\n", "2" },
long_bracket_nested_delimiters => { "print([=[hello]=])\n", "hello" },
escape_hex_two_digits => { "print(\"\\x41\")\n", "A" },
escape_unicode_curly_braces => { "print(\"\\u{61}\")\n", "a" },
userdata_type_tag_when_absent => { "print(type(io.stdin) == \"userdata\" or type(io.stdin) == \"nil\")\n", "true" },
integer_with_underscore_separators => { "print(1_000 + 2_000)\n", "3000" },
float_with_leading_zero => { "print(0.5 * 2)\n", "1" },
hex_in_expression => { "print(0x10 + 1)\n", "17" },
escape_single_quote_inside_double => { "print(\"it's\")\n", "it's" },
long_string_with_nested_brackets => { "print([==[a]=b]==])\n", "a]=b" },
nil_in_table_field_read => { "local t = {x = nil}\nprint(tostring(t.x))\n", "nil" },
true_and_false_in_expression => { "print(true and false or true)\n", "true" },
string_with_utf8_multibyte_char => { "print(#\"λ\")\n", "2" },
number_division_literal => { "print(10 / 4)\n", "2.5" },
hex_float_literal => {
    "print(0x1p4)\n",
    "16.0"
},
negative_hex_literal => {
    "print(-0xFF)\n",
    "-255"
},
octal_escape_sequence_in_string => {
    "print(\"\\065\")\n",
    "A"
},
long_string_ignores_leading_newline => {
    "local s = [[\nhello]]\nprint(s)\n",
    "hello"
},
scientific_notation_negative_exponent => {
    "print(1e-2)\n",
    "0.01"
},
empty_string_has_length_zero => {
    "print(#'')\n",
    "0"
},
boolean_true_converts_to_string_via_tostring => {
    "print(tostring(true) .. ',' .. tostring(false))\n",
    "true,false"
},
hex_float_with_fractional_part => {
    "print(0x1.8p1)\n",
    "3.0"
} }
