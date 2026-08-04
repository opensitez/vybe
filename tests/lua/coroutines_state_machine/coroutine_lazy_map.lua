-- vybe-test: lua/coroutines_state_machine/coroutine_lazy_map
-- origin: languages/lua/tests/lua/test_coroutines_state_machine.rs

local __w1 = "1,4,9,16"
local __i = 0

local function lazy_map(gen, f)
  return coroutine.wrap(function()
    for v in gen do coroutine.yield(f(v)) end
  end)
end
local function range(n)
  return coroutine.wrap(function()
    for i=1,n do coroutine.yield(i) end
  end)
end
local results = {}
for v in lazy_map(range(4), function(x) return x*x end) do
  results[#results+1] = v
end
do local __t = tostring(table.concat(results, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
