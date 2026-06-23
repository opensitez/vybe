use super::helpers::{compile_ok, parse_ok, run_lua_one};

// ── Literals ────────────────────────────────────────────────────────────────

#[test]
fn nil_literal() {
    parse_ok("local x = nil\n");
}

#[test]
fn boolean_literals() {
    parse_ok("local t = true\nlocal f = false\n");
}

#[test]
fn integer_literal() {
    parse_ok("local n = 42\n");
}

#[test]
fn float_literal() {
    parse_ok("local n = 3.14\n");
}

#[test]
fn string_literal() {
    parse_ok("local s = \"hello\"\n");
}

// ── Locals ──────────────────────────────────────────────────────────────────

#[test]
fn local_without_init() {
    parse_ok("local x\n");
}

#[test]
fn multiple_locals() {
    parse_ok("local a, b = 1, 2\n");
}

// ── print builtin ───────────────────────────────────────────────────────────

#[test]
fn print_integer() {
    let out = run_lua_one("print(42)\n");
    assert_eq!(out, "42");
}

#[test]
fn print_string() {
    let out = run_lua_one("print(\"hello\")\n");
    assert_eq!(out, "hello");
}

#[test]
fn print_concatenation() {
    let out = run_lua_one("print(\"a\" .. \"b\")\n");
    assert_eq!(out, "ab");
}

#[test]
fn print_nil() {
    let out = run_lua_one("print(nil)\n");
    assert_eq!(out, "nil");
}

#[test]
fn print_multiple_values_space_separated() {
    let out = run_lua_one("print(1, 2, 3)\n");
    assert_eq!(out, "1 2 3");
}

#[test]
fn multiple_assignment_swaps_via_temporaries() {
    let src = "local a, b = 1, 2\na, b = b, a\nprint(a .. \",\" .. b)\n";
    let out = run_lua_one(src);
    assert_eq!(out, "2,1");
}

#[test]
fn semicolon_separates_statements() {
    let out = run_lua_one("local x = 1; x = x + 1; print(x)\n");
    assert_eq!(out, "2");
}

// ── Functions ───────────────────────────────────────────────────────────────

#[test]
fn function_decl_compiles() {
    compile_ok("function add(a, b)\n  return a + b\nend\n");
}

#[test]
fn function_call_returns_sum() {
    let src = "function add(a, b)\n  return a + b\nend\nprint(add(2, 3))\n";
    let out = run_lua_one(src);
    assert_eq!(out, "5");
}

// ── Core daily usage: print, locals, types (Lua 5.x §2–3) ─────────────────

lua_print! {
    print_boolean_true => { "print(true)\n", "true" },
    print_boolean_false => { "print(false)\n", "false" },
    print_type_of_integer => { "print(type(42))\n", "number" },
    print_type_of_string => { "print(type(\"hi\"))\n", "string" },
    print_type_of_table => { "print(type({}))\n", "table" },
    local_reassignment_increment => {
        "local n = 1\nn = n + 1\nprint(n)\n",
        "2"
    },
    unset_local_is_nil => {
        "local x\nprint(tostring(x))\n",
        "nil"
    },
    if_local_then_branch_runs => {
        "local ok = true\nif ok then print(\"yes\") end\n",
        "yes"
    },
    global_name_readable_after_assignment => {
        "score = 10\nprint(score)\n",
        "10"
    },
    string_in_local_concatenated => {
        "local greeting = \"hello\"\nprint(greeting .. \" world\")\n",
        "hello world"
    },
    table_literal_in_local_indexed => {
        "local t = {\"a\", \"b\"}\nprint(t[1] .. t[2])\n",
        "ab"
    },
    call_function_stored_in_local => {
        "local function add(a, b) return a + b end\nprint(add(1, 2))\n",
        "3"
    },
    numeric_for_is_common_loop_form => {
        "local s = 0\nfor i = 1, 3 do s = s + i end\nprint(s)\n",
        "6"
    },
    while_loop_until_condition => {
        "local n = 1\nwhile n < 4 do n = n + 1 end\nprint(n)\n",
        "4"
    },
    repeat_until_common_form => {
        "local n = 0\nrepeat n = n + 1 until n == 2\nprint(n)\n",
        "2"
    },
    local_string_and_number_together => {
        "local name, age = \"ada\", 10\nprint(name .. age)\n",
        "ada10"
    },
    print_result_of_arithmetic_in_locals => {
        "local x = 2 * 3\nprint(x)\n",
        "6"
    },
    compare_locals_in_if_condition => {
        "local a, b = 3, 5\nif a < b then print(\"ok\") end\n",
        "ok"
    },
    nested_locals_in_do_block => {
        "local x = 1\n do local y = 2 x = x + y end\n print(x)\n",
        "3"
    },
    function_returns_to_print_directly => {
        "function two() return 2 end\nprint(two())\n",
        "2"
    },
    use_not_on_local_boolean => {
        "local f = false\nprint(not f)\n",
        "true"
    },
    concatenate_local_numbers_as_strings => {
        "local a, b = 1, 2\nprint(a .. b)\n",
        "12"
    },
    assign_global_then_read_in_print => {
        "version = 54\nprint(version)\n",
        "54"
    },
    if_else_picks_branch_on_local => {
        "local flag = false\nif flag then print(0) else print(1) end\n",
        "1"
    },
}
