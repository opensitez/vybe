//! Coroutine as state machine and lazy generator (Lua 5.x §2.6)

lua_print! {
    coroutine_state_transition => {
        "local function sm()\n  coroutine.yield(\"idle\")\n  coroutine.yield(\"running\")\n  coroutine.yield(\"stopped\")\nend\nlocal co = coroutine.create(sm)\nlocal states = {}\nfor _ = 1, 3 do\n  local _, s = coroutine.resume(co)\n  states[#states+1] = s\nend\nprint(table.concat(states, \",\"))\n",
        "idle,running,stopped"
    },
    coroutine_acc => {
        "local co = coroutine.create(function()\n  local acc = 0\n  while true do\n    local n = coroutine.yield(acc)\n    if n == nil then return acc end\n    acc = acc + n\n  end\nend)\ncoroutine.resume(co)  -- start\nfor _, v in ipairs({1,2,3,4,5}) do\n  coroutine.resume(co, v)\nend\nlocal _, total = coroutine.resume(co, nil)\nprint(total)\n",
        "15"
    },
    coroutine_fibonacci => {
        "local function fib_gen()\n  local a, b = 0, 1\n  while true do\n    coroutine.yield(a)\n    a, b = b, a + b\n  end\nend\nlocal co = coroutine.create(fib_gen)\nlocal nums = {}\nfor _ = 1, 7 do\n  local _, v = coroutine.resume(co)\n  nums[#nums+1] = v\nend\nprint(table.concat(nums, \",\"))\n",
        "0,1,1,2,3,5,8"
    },
    coroutine_steps_count => {
        "local steps = 0\nlocal co = coroutine.create(function()\n  for i = 1, 5 do\n    steps = steps + 1\n    coroutine.yield()\n  end\nend)\nwhile coroutine.status(co) ~= \"dead\" do\n  coroutine.resume(co)\nend\nprint(steps)\n",
        "5"
    },
    coroutine_lazy_map => {
        "local function lazy_map(gen, f)\n  return coroutine.wrap(function()\n    for v in gen do coroutine.yield(f(v)) end\n  end)\nend\nlocal function range(n)\n  return coroutine.wrap(function()\n    for i=1,n do coroutine.yield(i) end\n  end)\nend\nlocal results = {}\nfor v in lazy_map(range(4), function(x) return x*x end) do\n  results[#results+1] = v\nend\nprint(table.concat(results, \",\"))\n",
        "1,4,9,16"
    },
}
