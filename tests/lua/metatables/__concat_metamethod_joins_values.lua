-- vybe-test: lua/metatables/__concat_metamethod_joins_values
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "ab"
local __i = 0

local mt = {__concat = function(a,b) return a.v .. b.v end}
local a = setmetatable({v="a"}, mt)
local b = setmetatable({v="b"}, mt)
do local __t = tostring(a .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
