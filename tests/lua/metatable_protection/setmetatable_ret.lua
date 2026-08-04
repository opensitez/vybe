-- vybe-test: lua/metatable_protection/setmetatable_ret
-- origin: languages/lua/tests/lua/test_metatable_protection.rs

local __w1 = "true"
local __i = 0

local t = {}
local r = setmetatable(t, {})
do local __t = tostring(r == t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
