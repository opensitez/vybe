-- vybe-test: lua/errors/error_with_table_object
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "true"
local __i = 0

local err_obj = {code = 404}
local ok, caught = pcall(function() error(err_obj) end)
do local __t = tostring(ok == false and caught == err_obj and caught.code == 404); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
