-- vybe-test: lua/scoping/local_in_loop_body_recreated_each_iteration
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "30"
local __i = 0

local t = {}
for i = 1, 2 do
  local v = i * 10
  t[i] = v
end
do local __t = tostring(t[1] + t[2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
