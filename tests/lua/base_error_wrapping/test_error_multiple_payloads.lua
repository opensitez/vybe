-- vybe-test: lua/base_error_wrapping/test_error_multiple_payloads
-- origin: languages/lua/tests/lua/test_base_error_wrapping.rs

local __w1 = "true"
local __i = 0

local function f() return 1 end
local ok, err = pcall(function() local first = f(); error(first) end)
do local __t = tostring(ok == false and type(err) == "number"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
