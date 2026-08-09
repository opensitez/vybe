//! Complex upvalues capturing, variable lifetime, and shadow variables (Lua 5.x §3.5)

lua_print! {
lexical_three_levels => {
    "local x = 10\nlocal function f1()\n  local y = 20\n  return function()\n    local z = 30\n    return x + y + z\n  end\nend\nprint(f1()())\n",
    "60"
},
lexical_param_shadow => {
    "local x = 1\nlocal function f(x)\n  return function() return x end\nend\nprint(f(99)())\n",
    "99"
},
lexical_for_loop_capture => {
    "local fns = {}\nfor i = 1, 3 do\n  fns[i] = function() return i end\nend\nprint(fns[1]() .. fns[2]() .. fns[3]())\n",
    "123"
},
lexical_shared_mutation => {
    "local x = 0\nlocal f1 = function() x = x + 1; return x end\nlocal f2 = function() x = x + 10; return x end\nprint(f1() .. \",\" .. f2() .. \",\" .. f1())\n",
    "1,11,12"
},
lexical_shadow_block => {
    "local x = 5\nif true then\n  local x = 10\n  print(x)\nend\nprint(x)\n",
    "10\n5"
} }
