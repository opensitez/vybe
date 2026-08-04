-- vybe-test: lua/coroutine_patterns/producer_consumer
-- origin: languages/lua/tests/lua/test_coroutine_patterns.rs

local __w1 = "a,b,c"
local __i = 0

local function producer()
  local items = {"a", "b", "c"}
  for _, v in ipairs(items) do
    coroutine.yield(v)
  end
end
local co = coroutine.create(producer)
local results = {}
while true do
  local ok, v = coroutine.resume(co)
  if not ok or v == nil then break end
  results[#results+1] = v
end
do local __t = tostring(table.concat(results, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
