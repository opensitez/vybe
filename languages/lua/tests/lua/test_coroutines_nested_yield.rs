//! Yielding values through nested functions inside coroutines (Lua 5.x §2.6)

lua_print! {
nested_function_yield => {
    "local function inner() coroutine.yield(\"inner\") end\nlocal function outer() inner(); return \"outer\" end\nlocal co = coroutine.create(outer)\nlocal _, v1 = coroutine.resume(co)\nlocal _, v2 = coroutine.resume(co)\nprint(v1 .. \",\" .. v2)\n",
    "inner,outer"
},
yield_inside_pcall => {
    "local co = coroutine.create(function()\n  local ok, err = pcall(function()\n    coroutine.yield(\"yielded\")\n  end)\n  return ok, err\nend)\nlocal _, val = coroutine.resume(co)\nprint(val)\n",
    "yielded"
},
coroutine_status_during_yield => {
    "local co\nco = coroutine.create(function()\n  coroutine.yield(coroutine.status(co))\nend)\nlocal _, status = coroutine.resume(co)\nprint(status)\n",
    "running"
},
multiple_nested_yields => {
    "local function step()\n  local x = coroutine.yield(\"need_x\")\n  local y = coroutine.yield(\"need_y\")\n  return x + y\nend\nlocal co = coroutine.create(step)\ncoroutine.resume(co)\ncoroutine.resume(co, 5)\nlocal _, sum = coroutine.resume(co, 10)\nprint(sum)\n",
    "15"
},
yield_inside_for_iterator => {
    "local co = coroutine.create(function()\n  for i = 1, 3 do coroutine.yield(i) end\n  return 4\nend)\nlocal r = \"\"\nwhile coroutine.status(co) ~= \"dead\" do\n  local ok, val = coroutine.resume(co)\n  if ok and val then r = r .. val end\nend\nprint(r)\n",
    "1234"
} }
