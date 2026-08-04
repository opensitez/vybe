-- vybe-test: lua/type_checks/tostring_on_table_uses_metamethod
-- origin: languages/lua/tests/lua/test_type_checks.rs

local __w1 = "T"
local __i = 0

local t=setmetatable({}, {__tostring=function() return "T" end})
do local __t = tostring(tostring(t)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
