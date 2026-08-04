-- vybe-test: lua/metamethods/metamethod_concat_joins_custom_values
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "xy"
local __i = 0

local mt = {__concat = function(a, b) return a.s .. b.s end}
local a = setmetatable({s = "x"}, mt)
local b = setmetatable({s = "y"}, mt)
do local __t = tostring(a .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
