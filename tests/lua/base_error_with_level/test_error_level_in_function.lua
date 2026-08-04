-- vybe-test: lua/base_error_with_level/test_error_level_in_function
-- origin: languages/lua/tests/lua/test_base_error_with_level.rs

local __w1 = "true"
local __i = 0

local function f()
  error("inner", 2)
end
local ok, err = pcall(f)
do local __t = tostring(string.find(err, "inner") ~= nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
