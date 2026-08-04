-- vybe-test: lua/metatables_call/test_call_string
-- origin: languages/lua/tests/lua/test_metatables_call.rs

local __w1 = "hello world"
local __i = 0

debug.setmetatable('', {__call=function(s, a) return s..a end}); do local __t = tostring(('hello ')('world')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
