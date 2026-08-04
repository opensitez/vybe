-- vybe-test: lua/base_getmetatable_presence/test_getmetatable_presence_captured
-- origin: languages/lua/tests/lua/test_base_getmetatable_presence.rs

local __w1 = "true"
local __i = 0

local t = {}
local mt = {__tostring = function() return "ok" end}
setmetatable(t, mt)
do local __t = tostring(getmetatable(t) == mt); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
