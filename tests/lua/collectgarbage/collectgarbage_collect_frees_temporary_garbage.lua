-- vybe-test: lua/collectgarbage/collectgarbage_collect_frees_temporary_garbage
-- origin: languages/lua/tests/lua/test_collectgarbage.rs

local __w1 = "true"
local __i = 0

local before = collectgarbage("count")
local function make_garbage()
  local t = {}
  for i=1,1000 do t[i] = {x=i} end
end
make_garbage()
collectgarbage("collect")
local after = collectgarbage("count")
do local __t = tostring(after - before < 100); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
