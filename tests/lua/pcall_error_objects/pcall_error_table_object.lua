-- vybe-test: lua/pcall_error_objects/pcall_error_table_object
-- origin: languages/lua/tests/lua/test_pcall_error_objects.rs

local __w1 = "false 42"
local __i = 0

local ok, err = pcall(function() error({code=42}) end)
do local __t = tostring(ok) .. "\t" .. tostring(err.code); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
