-- vybe-test: lua/pcall_error_objects/pcall_table_fields
-- origin: languages/lua/tests/lua/test_pcall_error_objects.rs

local __w1 = "3"
local __i = 0

local ok, e = pcall(function() error({a=1, b=2}) end)
do local __t = tostring(e.a + e.b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
