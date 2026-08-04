-- vybe-test: lua/debug_introspection/test_debug_getinfo_fields
-- origin: languages/lua/tests/lua/test_debug_introspection.rs

local __w1 = "string number string"
local __i = 0

local info = debug.getinfo(1, 'nlS'); do local __t = tostring(type(info.name)..' '..type(info.currentline)..' '..type(info.short_src)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
