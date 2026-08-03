//! Local vs global scope — Lua 5.x manual §2.6.

lua_print! {
    local_shadows_global_read => {
        "x = 1\nlocal x = 2\nprint(x)\n",
        "2"
    },
    assignment_after_local_updates_local => {
        "local n = 1\nn = n + 4\nprint(n)\n",
        "5"
    },
    multiple_locals_in_one_statement => {
        "local a, b = 3, 4\nprint(a + b)\n",
        "7"
    },
    local_in_block_not_visible_outside => {
        "do local x = 99 end\nprint(x)\n",
        "nil"
    },
    nested_do_block_inner_local_shadows_outer => {
        "local v = 1\n do\n  local v = 2\n  print(v)\n end\n",
        "2"
    },
    outer_local_visible_after_do_block_ends => {
        "local v = 1\n do local v = 2 end\n print(v)\n",
        "1"
    },
    function_parameter_shadows_outer_local_in_body => {
        "local n = 1\nfunction f(n) return n end\nprint(f(9))\n",
        "9"
    },
    outer_local_unchanged_when_parameter_shadows => {
        "local n = 1\nfunction f(n) end\nprint(n)\n",
        "1"
    },
    local_function_forward_reference_sugar => {
        "local function fact(n)\n  if n <= 1 then return 1 end\n  return n * fact(n - 1)\nend\nprint(fact(4))\n",
        "24"
    },
    upvalue_shared_between_nested_closures => {
        "local n = 0\nlocal function inc() n = n + 1 end\nlocal function read() return n end\ninc()\nprint(read())\n",
        "1"
    },
    chunk_local_not_visible_in_later_function_if_not_upvalue => {
        "local z = 5\nfunction g() return z end\nz = 6\nprint(g())\n",
        "6"
    },
    global_written_inside_function_visible_outside => {
        "function setk() key = 9 end\nsetk()\nprint(key)\n",
        "9"
    },
    local_in_loop_body_recreated_each_iteration => {
        "local t = {}\nfor i = 1, 2 do\n  local v = i * 10\n  t[i] = v\nend\nprint(t[1] + t[2])\n",
        "30"
    },
    do_block_limits_local_scope => {
        "local x = 1\n do local x = 5 end\n print(x)\n",
        "1"
    },
    function_name_in_outer_scope_not_local => {
        "function h() return 1 end\nprint(h())\n",
        "1"
    },
    scoping_upvalues_independent_in_separate_closure_instances => {
        "local function make_counter()\n  local count = 0\n  return function() count = count + 1; return count end\nend\nlocal c1 = make_counter()\nlocal c2 = make_counter()\nc1(); c1()\nprint(c1() .. \" \" .. c2())\n",
        "3 1"
    },
    scoping_upvalue_mutated_after_closure_creation_is_visible => {
        "local x = 10\nlocal function get_x() return x end\nx = 20\nprint(get_x())\n",
        "20"
    },
    scoping_multiple_locals_declared_with_same_name_in_single_line => {
        "local x, x = 100, 200\nprint(x)\n",
        "200"
    },
    scoping_repeat_until_condition_can_access_locals_in_loop_body => {
        "local count = 0\nrepeat\n  local x = 10\n  count = count + x\nuntil x == 10\nprint(count)\n",
        "10"
    },
    scoping_generic_for_loop_variables_are_local_to_body => {
        "local k = \"outer\"\nlocal t = {key = \"val\"}\nfor k, v in pairs(t) do end\nprint(k)\n",
        "outer"
    },
    scoping_numeric_for_loop_control_variable_is_read_only => {
        "local sum = 0\nfor i = 1, 3 do\n  sum = sum + i\n  i = 100\nend\nprint(sum)\n",
        "6"
    },
    scoping_nested_functions_share_upvalue_chain => {
        "local val = 5\nlocal function f1()\n  local function f2()\n    local function f3()\n      val = val + 10\n    end\n    f3()\n  end\n  f2()\nend\nf1()\nprint(val)\n",
        "15"
    },
    scoping_custom_env_redirects_global_writes => {
        "local env = {}\nlocal function run_in_env()\n  local _ENV = env\n  global_var = 42\nend\nrun_in_env()\nprint(env.global_var)\n",
        "42"
    },
    scoping_closure_created_inside_loop_captures_correct_variable_instance => {
        "local funcs = {}\nfor i = 1, 2 do\n  local val = i\n  funcs[i] = function() return val end\nend\nprint(funcs[1]() .. \" \" .. funcs[2]())\n",
        "1 2"
    },
    local_in_then_block_not_visible_after_end => {
        "if true then\n  local secret = 'yes'\nend\nprint(tostring(secret))\n",
        "nil"
    },
    three_levels_of_shadowing_innermost_wins => {
        "local x = 'outer'\ndo\n  local x = 'middle'\n  do\n    local x = 'inner'\n    print(x)\n  end\n  print(x)\nend\nprint(x)\n",
        "inner\nmiddle\nouter"
    },
    upvalue_from_function_param_survives_return => {
        "local function make(n)\n  return function() return n end\nend\nlocal f = make(55)\nprint(f())\n",
        "55"
    },
    for_generic_control_vars_are_local_to_body => {
        "local i = 'outer'\nfor i, v in ipairs({'a', 'b'}) do end\nprint(i)\n",
        "outer"
    },
    local_declaration_hides_same_name_global => {
        "g_conflict = 'global'\ndo\n  local g_conflict = 'local'\n  print(g_conflict)\nend\nprint(g_conflict)\n",
        "local\nglobal"
    },
    local_not_initialized_starts_as_nil => {
        "local uninit\nprint(tostring(uninit))\n",
        "nil"
    },
    upvalue_written_before_closure_created_is_seen => {
        "local x\nx = 10\nlocal f = function() return x end\nprint(f())\n",
        "10"
    },
    multiple_upvalues_from_same_scope_in_single_closure => {
        "local a, b, c = 1, 2, 3\nlocal f = function() return a + b + c end\nprint(f())\n",
        "6"
    } }
