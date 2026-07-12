//! Coroutine producers and pipelines — real-world coroutine use patterns (Lua 5.x §2.6)

lua_print! {
    producer_consumer => {
        "local function producer()\n  local items = {\"a\", \"b\", \"c\"}\n  for _, v in ipairs(items) do\n    coroutine.yield(v)\n  end\nend\nlocal co = coroutine.create(producer)\nlocal results = {}\nwhile true do\n  local ok, v = coroutine.resume(co)\n  if not ok or v == nil then break end\n  results[#results+1] = v\nend\nprint(table.concat(results, \",\"))\n",
        "a,b,c"
    },
    coroutine_lazy_range => {
        "local function range(n)\n  return coroutine.wrap(function()\n    for i = 1, n do coroutine.yield(i) end\n  end)\nend\nlocal t = {}\nfor v in range(5) do t[#t+1] = v end\nprint(t[3])\n",
        "3"
    },
    coroutine_pipeline => {
        "local function filter(gen, pred)\n  return coroutine.wrap(function()\n    for v in gen do\n      if pred(v) then coroutine.yield(v) end\n    end\n  end)\nend\nlocal function nums(n)\n  return coroutine.wrap(function()\n    for i = 1, n do coroutine.yield(i) end\n  end)\nend\nlocal evens = filter(nums(10), function(v) return v % 2 == 0 end)\nlocal sum = 0\nfor v in evens do sum = sum + v end\nprint(sum)\n",
        "30"
    },
    coroutine_accumulate => {
        "local co = coroutine.create(function()\n  for i = 1, 4 do coroutine.yield(i * i) end\nend)\nlocal sum = 0\nfor _ = 1, 4 do\n  local _, v = coroutine.resume(co)\n  sum = sum + v\nend\nprint(sum)\n",
        "30"
    },
    coroutine_communication => {
        "local co = coroutine.create(function(start)\n  local acc = start\n  while true do\n    local n = coroutine.yield(acc)\n    if n == nil then break end\n    acc = acc + n\n  end\n  return acc\nend)\nlocal _, v1 = coroutine.resume(co, 10)\nlocal _, v2 = coroutine.resume(co, 5)\nlocal _, v3 = coroutine.resume(co, 3)\nprint(v1, v2, v3)\n",
        "10\t15\t18"
    },
    coroutine_resume_dead => {
        "local co = coroutine.create(function() return 1 end)\ncoroutine.resume(co)\nlocal ok = coroutine.resume(co)\nprint(ok)\n",
        "false"
    },
    coroutine_wrap_error => {
        "local f = coroutine.wrap(function() error(\"crash\") end)\nlocal ok = pcall(f)\nprint(ok)\n",
        "false"
    },
    coroutine_status_normal => {
        "local main_status\nlocal co2 = coroutine.create(function()\n  main_status = coroutine.status(coroutine.running())\nend)\ncoroutine.resume(co2)\nprint(main_status)\n",
        "running"
    },
}
