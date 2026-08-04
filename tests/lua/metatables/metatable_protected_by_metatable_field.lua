-- vybe-test: lua/metatables/metatable_protected_by_metatable_field
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "false"
local __i = 0

local mt = {__metatable = "locked"}
local t = setmetatable({}, mt)
local ok = pcall(function() setmetatable(t, {}) end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
