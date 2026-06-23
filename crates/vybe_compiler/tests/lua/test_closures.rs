//! Closures and upvalues — nested scopes (Lua 5.x manual §3.4.11).

lua_print! {
    closure_captures_parameter_value => {
        "function make(x)\n  return function() return x end\nend\nprint(make(3)())\n",
        "3"
    },
    closure_sees_later_mutation_of_upvalue => {
        "local n=1\nlocal f=function() return n end\nn=2\nprint(f())\n",
        "2"
    },
    nested_closure_shares_outer_upvalue => {
        "local n=0\nlocal function outer()\n  local function inner() n=n+1 end\n  return inner\nend\nlocal inc=outer()\ninc()\nprint(n)\n",
        "1"
    },
    each_closure_has_distinct_upvalue_binding => {
        "local function make()\n  local n=0\n  return function() n=n+1 return n end\nend\nlocal a=make()\nlocal b=make()\na()\nprint(b())\n",
        "1"
    },
    closure_over_loop_variable_requires_iife_pattern => {
        "local fns={}\nfor i=1,2 do\n  fns[i]=function() return i end\nend\nprint(fns[1]()+fns[2]())\n",
        "3"
    },
    returning_closure_preserves_state => {
        "function counter()\n  local n=0\n  return function() n=n+1 return n end\nend\nlocal c=counter()\nc()\nprint(c())\n",
        "2"
    },
    upvalue_not_visible_to_sibling_inner_functions => {
        "local n=0\nlocal function a() n=1 end\nlocal function b() return n end\na()\nprint(b())\n",
        "1"
    },
    global_lookup_when_not_local => {
        "gval=5\nlocal function f() return gval end\nprint(f())\n",
        "5"
    },
    local_shadows_global_in_closure_body => {
        "x=1\nlocal function f()\n  local x=2\n  return function() return x end\nend\nprint(f()())\n",
        "2"
    },
    deeply_nested_closure_reads_outermost => {
        "local function a()\n  local v=1\n  return function()\n    return function() return v end\n  end\nend\nprint(a()()())\n",
        "1"
    },
    closure_passed_as_callback_to_iterating_helper => {
        "local function map(t, f)\n  local out = {}\n  for i, v in ipairs(t) do out[i] = f(v) end\n  return out\nend\nprint(map({1,2}, function(x) return x * 10 end)[2])\n",
        "20"
    },
    function_returning_function_factory => {
        "local function make_adder(n)\n  return function(x) return x + n end\nend\nprint(make_adder(5)(3))\n",
        "8"
    },
    stored_function_reference_in_table => {
        "local api = { run = function() return 9 end }\nprint(api.run())\n",
        "9"
    },
    function_argument_is_another_function => {
        "local function apply(f) return f() end\nprint(apply(function() return \"ok\" end))\n",
        "ok"
    },
}
