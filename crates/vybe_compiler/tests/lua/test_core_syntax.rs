//! Core syntax smoke — parse/compile checks for everyday constructs (Lua 5.x §3).
//! Runtime gaps are exercised elsewhere; these assert the frontend accepts common forms.

use super::helpers::{compile_ok, parse_ok};

#[test]
fn parse_function_with_params_and_return() {
    parse_ok("function f(a, b)\n  return a + b\nend\n");
}

#[test]
fn parse_local_function_sugar() {
    parse_ok("local function g() end\n");
}

#[test]
fn parse_if_elseif_else_end() {
    parse_ok("if true then x=1 elseif false then x=2 else x=3 end\n");
}

#[test]
fn parse_while_do_end() {
    parse_ok("while true do break end\n");
}

#[test]
fn parse_repeat_until() {
    parse_ok("repeat x=1 until true\n");
}

#[test]
fn parse_numeric_for_loop() {
    parse_ok("for i = 1, 10 do print(i) end\n");
}

#[test]
fn parse_generic_for_in() {
    parse_ok("for k, v in pairs(t) do print(k, v) end\n");
}

#[test]
fn parse_table_constructor_list() {
    parse_ok("local t = {1, 2, 3}\n");
}

#[test]
fn parse_table_constructor_record() {
    parse_ok("local t = {a = 1, b = 2}\n");
}

#[test]
fn parse_table_index_and_field() {
    parse_ok("local t = {}\nt[1] = 1\nt.x = 2\n");
}

#[test]
fn parse_string_concat_and_length() {
    parse_ok("local s = \"a\" .. \"b\"\nlocal n = #s\n");
}

#[test]
fn parse_operator_precedence_expression() {
    parse_ok("local x = 1 + 2 * 3 ^ 2\n");
}

#[test]
fn compile_simple_print_script() {
    compile_ok("print(\"ok\")\n");
}

#[test]
fn compile_local_assignment_chain() {
    compile_ok("local a, b, c = 1, 2, 3\nprint(a + b + c)\n");
}

#[test]
fn compile_function_call_in_expression() {
    compile_ok("function id(x) return x end\nprint(id(9))\n");
}
