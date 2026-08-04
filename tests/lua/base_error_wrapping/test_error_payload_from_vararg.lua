-- vybe-test: lua/base_error_wrapping/test_error_payload_from_vararg
-- origin: languages/lua/tests/lua/test_base_error_wrapping.rs

local __w1 = "true"
local __i = 0

local function f() return 2 end
local ok, err = pcall(function()
  local b = f()
  error(b)
end)
do local __t = tostring(type(err) == "number"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
