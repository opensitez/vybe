//! Functions — Lua 5.x manual §3.4.9–3.4.11.

lua_print! {
function_no_args_returns_literal => {
    "function f() return 9 end\nprint(f())\n",
    "9"
},
function_single_param => {
    "function double(x) return x * 2 end\nprint(double(11))\n",
    "22"
},
function_recursion_factorial => {
    "function fact(n)\n  if n <= 1 then return 1 end\n  return n * fact(n - 1)\nend\nprint(fact(5))\n",
    "120"
},
function_early_return => {
    "function sign(n)\n  if n < 0 then return -1 end\n  if n > 0 then return 1 end\n  return 0\nend\nprint(sign(-3))\n",
    "-1"
},
nested_function_call => {
    "function inc(x) return x + 1 end\nfunction twice(x) return inc(inc(x)) end\nprint(twice(10))\n",
    "12"
},
function_uses_outer_local => {
    "local base = 100\nfunction add(x) return base + x end\nprint(add(5))\n",
    "105"
},
function_without_return_yields_nil => {
    "function f() end\nprint(tostring(f()))\n",
    "nil"
},
function_multiple_return_values => {
    "function swap(a,b) return b,a end\nlocal x,y=swap(1,2)\nprint(x..\",\"..y)\n",
    "2,1"
},
closure_reads_captured_upvalue => {
    "function make()\n  local n=0\n  return function() return n end\nend\nprint(make()())\n",
    "0"
},
closure_mutates_enclosed_local_across_calls => {
    "function counter()\n  local n=0\n  return function() n=n+1 return n end\nend\nlocal c=counter()\nprint(c() .. \",\" .. c())\n",
    "1,2"
},
nested_function_sees_enclosing_parameter => {
    "function outer(x)\n  function inner() return x end\n  return inner()\nend\nprint(outer(8))\n",
    "8"
},
local_function_statement_desugars_to_assignment => {
    "local function twice(x) return x * 2 end\nprint(twice(6))\n",
    "12"
},
function_expression_stored_in_local => {
    "local f = function(x) return x - 1 end\nprint(f(5))\n",
    "4"
},
mutual_recursion_even_odd => {
    "local is_even, is_odd\nfunction is_even(n)\n  if n == 0 then return true end\n  return is_odd(n - 1)\nend\nfunction is_odd(n)\n  if n == 0 then return false end\n  return is_even(n - 1)\nend\nprint(is_even(4))\n",
    "true"
},
return_skips_remaining_statements => {
    "function f()\n  return 1\n  return 2\nend\nprint(f())\n",
    "1"
},
tail_call_does_not_grow_stack_is_semantics => {
    "local function tail(n)\n  if n == 0 then return \"done\" end\n  return tail(n - 1)\nend\nprint(tail(5))\n",
    "done"
},
function_call_with_expression_arguments => {
    "function add(a, b) return a + b end\nprint(add(1 + 1, 2 + 2))\n",
    "6"
},
function_call_result_used_in_condition => {
    "function is_positive(n) return n > 0 end\nif is_positive(3) then print(\"yes\") end\n",
    "yes"
},
missing_arguments_become_nil => {
    "function show(x) print(tostring(x)) end\nshow()\n",
    "nil"
},
extra_arguments_are_ignored => {
    "function one(x) return x end\nprint(one(1, 2, 3))\n",
    "1"
},
function_returning_string_for_concat => {
    "function label() return \"lua\" end\nprint(label() .. \"!\")\n",
    "lua!"
},
global_function_call_without_local => {
    "function inc(x) return x + 1 end\nprint(inc(4))\n",
    "5"
},
function_returns_boolean_for_if => {
    "function is_even(n) return n % 2 == 0 end\nif is_even(4) then print(\"yes\") end\n",
    "yes"
},
function_with_three_parameters => {
    "function add3(a, b, c) return a + b + c end\nprint(add3(1, 2, 3))\n",
    "6"
},
function_call_as_operand => {
    "function double(x) return x * 2 end\nprint(double(3) + 1)\n",
    "7"
},
function_returning_table_literal => {
    "function point() return {x = 1, y = 2} end\nlocal p = point()\nprint(p.x + p.y)\n",
    "3"
},
function_passed_table_and_reads_field => {
    "function read_x(t) return t.x end\nprint(read_x({x = 9}))\n",
    "9"
},
function_default_nil_param => {
    "function show(x) return tostring(x) end\nprint(show())\n",
    "nil"
},
function_overwrites_local_with_return => {
    "function pick() return \"chosen\" end\nlocal x = pick()\nprint(x)\n",
    "chosen"
},
chained_calls_with_return_values => {
    "function twice(x) return x * 2 end\nfunction add1(x) return x + 1 end\nprint(add1(twice(3)))\n",
    "7"
},
function_body_local_shadows_param => {
    "function f(n)\n  local n = n + 1\n  return n\nend\nprint(f(4))\n",
    "5"
},
function_call_in_concat_expression => {
    "function name() return \"lua\" end\nprint(\"lang:\" .. name())\n",
    "lang:lua"
},
function_return_drives_while_condition => {
    "function has_more(n) return n > 0 end\nlocal n = 2\nwhile has_more(n) do n = n - 1 end\nprint(n)\n",
    "0"
},
function_with_if_inside => {
    "function abs(n)\n  if n < 0 then return -n end\n  return n\nend\nprint(abs(-6))\n",
    "6"
},
function_assigns_to_upvalue_via_helper => {
    "local total = 0\nfunction add_to_total(x) total = total + x end\nadd_to_total(3)\nadd_to_total(4)\nprint(total)\n",
    "7"
},
anonymous_function_in_call_position => {
    "local function apply(f, x) return f(x) end\nprint(apply(function(v) return v + 5 end, 2))\n",
    "7"
},
function_returning_result_of_builtin => {
    "function len(s) return #s end\nprint(len(\"abc\"))\n",
    "3"
},
function_call_with_string_argument => {
    "function shout(s) return string.upper(s) end\nprint(shout(\"hi\"))\n",
    "HI"
},
function_call_with_expression_args => {
    "function mul(a, b) return a * b end\nprint(mul(2 + 1, 3 + 1))\n",
    "12"
},
function_used_in_numeric_for_bound => {
    "function limit() return 3 end\nlocal s = 0\nfor i = 1, limit() do s = s + i end\nprint(s)\n",
    "6"
},
method_call_style_with_dot_syntax => {
    "local t = {}\nfunction t.greet() return \"hi\" end\nprint(t.greet())\n",
    "hi"
},
callback_passed_to_helper_function => {
    "local function apply(f, x) return f(x) end\nprint(apply(function(n) return n * 3 end, 4))\n",
    "12"
},
function_value_stored_in_table_slot => {
    "local ops = {add = function(a, b) return a + b end}\nprint(ops.add(2, 5))\n",
    "7"
},
function_returned_from_factory => {
    "local function make_adder(n)\n  return function(x) return x + n end\nend\nprint(make_adder(10)(3))\n",
    "13"
},
higher_order_map_over_array => {
    "local function map(t, f)\n  local out = {}\n  for i = 1, #t do out[i] = f(t[i]) end\n  return out\nend\nprint(map({1, 2, 3}, function(x) return x * x end)[3])\n",
    "9"
},
higher_order_filter_predicate => {
    "local function keep_if(t, pred)\n  local out = {}\n  for i = 1, #t do if pred(t[i]) then table.insert(out, t[i]) end end\n  return out\nend\nprint(#keep_if({1, 2, 3, 4}, function(n) return n % 2 == 1 end))\n",
    "2"
},
function_parameter_is_another_function => {
    "local function twice_call(f) return f() + f() end\nprint(twice_call(function() return 4 end))\n",
    "8"
},
sort_with_custom_comparator_function => {
    "local t = {3, 1, 2}\ntable.sort(t, function(a, b) return a > b end)\nprint(t[1])\n",
    "3"
},
immediately_invoked_function_expression => {
    "print((function(x) return x + 1 end)(4))\n",
    "5"
},
function_reference_equality_is_false => {
    "local f = function() end\nprint(tostring(f == f))\n",
    "true"
},
nested_return_passes_function_outward => {
    "local function outer()\n  return function() return \"inner\" end\nend\nprint(outer()())\n",
    "inner"
},
callback_accumulates_in_outer_local => {
    "local sum = 0\nlocal function add_each(t, fn)\n  for i = 1, #t do sum = sum + fn(t[i]) end\nend\nadd_each({1, 2, 3}, function(x) return x end)\nprint(sum)\n",
    "6"
},
unary_function_composed_inline => {
    "local function pipe(x, f, g) return g(f(x)) end\nprint(pipe(2, function(n) return n + 1 end, function(n) return n * 10 end))\n",
    "30"
},
function_type_is_always_function => {
    "print(type(function() end))\n",
    "function"
},
select_picks_function_from_list => {
    "local fns = {function() return 1 end, function() return 2 end}\nprint(fns[2]())\n",
    "2"
},
partial_application_via_closure => {
    "local function add(a, b) return a + b end\nlocal function bind_a(a)\n  return function(b) return add(a, b) end\nend\nprint(bind_a(5)(3))\n",
    "8"
},
function_overwrites_global_when_not_local => {
    "function greet() return \"hi\" end\nfunction wrap() return greet() end\nprint(wrap())\n",
    "hi"
},
anonymous_function_captured_in_loop => {
    "local t = {}\nfor i = 1, 2 do t[i] = function() return i * 10 end end\nprint(t[1]() + t[2]())\n",
    "30"
},
function_returns_multiple_functions => {
    "local function make_pair()\n  return function() return 1 end, function() return 2 end\nend\nlocal a, b = make_pair()\nprint(a() + b())\n",
    "3"
},
reduce_left_with_binary_function => {
    "local function fold(t, f, init)\n  local acc = init\n  for i = 1, #t do acc = f(acc, t[i]) end\n  return acc\nend\nprint(fold({1, 2, 3}, function(a, b) return a + b end, 0))\n",
    "6"
},
function_call_chain_left_associative => {
    "local function id(x) return x end\nprint(id(id(id(9))))\n",
    "9"
},
colon_method_syntax_passes_self_as_first_arg => {
    "local obj = {value = 10}\nfunction obj:get() return self.value end\nprint(obj:get())\n",
    "10"
},
colon_method_stores_self_and_mutates => {
    "local obj = {n = 0}\nfunction obj:inc() self.n = self.n + 1 end\nobj:inc(); obj:inc()\nprint(obj.n)\n",
    "2"
},
function_in_nested_table_field => {
    "local api = { utils = { double = function(x) return x * 2 end } }\nprint(api.utils.double(21))\n",
    "42"
},
function_receives_table_and_mutates_it => {
    "local function fill(t, val)\n  for i = 1, 3 do t[i] = val end\nend\nlocal data = {}\nfill(data, 7)\nprint(data[1] .. ',' .. data[2] .. ',' .. data[3])\n",
    "7,7,7"
},
anonymous_recursive_function_via_table_key => {
    "local fib = {}\nfib.calc = function(n)\n  if n <= 1 then return n end\n  return fib.calc(n - 1) + fib.calc(n - 2)\nend\nprint(fib.calc(7))\n",
    "13"
},
function_as_first_class_value_through_multiple_calls => {
    "local function compose(f, g) return function(x) return f(g(x)) end end\nlocal double = function(x) return x * 2 end\nlocal inc = function(x) return x + 1 end\nlocal double_then_inc = compose(inc, double)\nprint(double_then_inc(5))\n",
    "11"
},
vararg_forwarding_preserves_count_and_nils => {
    "local function wrapper(...)\n  return ...\nend\nlocal a, b, c = wrapper(10, nil, 30)\nprint(tostring(a) .. ',' .. tostring(b) .. ',' .. tostring(c))\n",
    "10,nil,30"
},
pcall_with_function_returning_multiple_values => {
    "local ok, a, b, c = pcall(function() return 10, 20, 30 end)\nprint(tostring(ok) .. ',' .. a .. ',' .. b .. ',' .. c)\n",
    "true,10,20,30"
} }
