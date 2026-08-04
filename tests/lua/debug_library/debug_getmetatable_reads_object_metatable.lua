-- vybe-test: lua/debug_library/debug_getmetatable_reads_object_metatable
-- origin: languages/lua/tests/lua/test_debug_library.rs

local __w1 = "true"
local __i = 0

local m = {}
local t = setmetatable({}, m)
do local __t = tostring(debug.getmetatable(t) == m); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
