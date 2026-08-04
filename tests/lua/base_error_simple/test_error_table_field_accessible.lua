-- vybe-test: lua/base_error_simple/test_error_table_field_accessible
-- origin: languages/lua/tests/lua/test_base_error_simple.rs

local __w1 = "true"
local __i = 0

local payload = {code = 500}
local ok, err = pcall(function() error(payload) end)
do local __t = tostring(type(err) == "table" and err.code == 500); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
