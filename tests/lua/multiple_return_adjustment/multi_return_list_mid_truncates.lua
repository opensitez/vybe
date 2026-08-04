-- vybe-test: lua/multiple_return_adjustment/multi_return_list_mid_truncates
-- origin: languages/lua/tests/lua/test_multiple_return_adjustment.rs

local __w1 = "4,9,nil"
local __i = 0

local function f() return 4, 5 end
local a, b, c = f(), 9
do local __t = tostring(a .. "," .. b .. "," .. tostring(c)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
