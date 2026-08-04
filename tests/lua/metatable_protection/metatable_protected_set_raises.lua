-- vybe-test: lua/metatable_protection/metatable_protected_set_raises
-- origin: languages/lua/tests/lua/test_metatable_protection.rs

local __w1 = "false"
local __i = 0

local mt = {__metatable = "lock"}
local t = setmetatable({}, mt)
local ok = pcall(setmetatable, t, {})
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
