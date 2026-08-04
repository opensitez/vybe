-- vybe-test: lua/metatable_protection/metatable_shared
-- origin: languages/lua/tests/lua/test_metatable_protection.rs

local __w1 = "42\t42"
local __i = 0

local mt = {__index = function() return 99 end}
local a = setmetatable({}, mt)
local b = setmetatable({}, mt)
mt.__index = function() return 42 end
do local __t = tostring(a.x) .. "\t" .. tostring(b.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
