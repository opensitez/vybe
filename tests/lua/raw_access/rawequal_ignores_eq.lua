-- vybe-test: lua/raw_access/rawequal_ignores_eq
-- origin: languages/lua/tests/lua/test_raw_access.rs

local __w1 = "false"
local __i = 0

local mt = {__eq = function() return true end}
local a = setmetatable({}, mt)
local b = setmetatable({}, mt)
do local __t = tostring(rawequal(a, b)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
