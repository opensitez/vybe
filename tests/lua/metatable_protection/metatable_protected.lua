-- vybe-test: lua/metatable_protection/metatable_protected
-- origin: languages/lua/tests/lua/test_metatable_protection.rs

local __w1 = "protected"
local __i = 0

local mt = {__metatable = "protected"}
local t = setmetatable({}, mt)
do local __t = tostring(getmetatable(t)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
