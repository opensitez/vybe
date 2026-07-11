//! Logical operators (and, or, not) with truthiness and lazy evaluation (Lua 5.x §3.4.5)

lua_print! {
    logical_and_nil => {
        "print(nil and 5)\n",
        "nil"
    },
    logical_and_false => {
        "print(false and 5)\n",
        "false"
    },
    logical_and_truthy => {
        "print(5 and 10)\n",
        "10"
    },
    logical_or_truthy_num => {
        "print(5 or 10)\n",
        "5"
    },
    logical_or_truthy_str => {
        "print(\"hi\" or false)\n",
        "hi"
    },
    logical_or_falsy => {
        "print(nil or 42)\n",
        "42"
    },
    logical_not_num => {
        "print(not 5)\n",
        "false"
    },
    logical_not_nil => {
        "print(not nil)\n",
        "true"
    },
    logical_not_false => {
        "print(not false)\n",
        "true"
    },
    logical_not_table => {
        "print(not {})\n",
        "false"
    },
    logical_and_short_circuit => {
        "local called = false\nlocal function rhs() called = true; return true end\nlocal _ = false and rhs()\nprint(called)\n",
        "false"
    },
    logical_or_short_circuit => {
        "local called = false\nlocal function rhs() called = true; return true end\nlocal _ = true or rhs()\nprint(called)\n",
        "false"
    },
    logical_coalesce => {
        "local x = nil\nlocal y = 42\nprint(x or y or 100)\n",
        "42"
    },
    logical_double_not => {
        "print(not not 5)\n",
        "true"
    },
    logical_ternary_true => {
        "local condition = true\nlocal val = condition and \"yes\" or \"no\"\nprint(val)\n",
        "yes"
    },
    logical_ternary_false => {
        "local condition = false\nlocal val = condition and \"yes\" or \"no\"\nprint(val)\n",
        "no"
    },
}
