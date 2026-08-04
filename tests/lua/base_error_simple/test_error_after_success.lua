-- vybe-test: lua/base_error_simple/test_error_after_success
-- origin: languages/lua/tests/lua/test_base_error_simple.rs

local __w1 = "7"
local __i = 0

local function check()
  local ok, err = pcall(function() return 1 end)
  if ok then return 7 end
  return err
end
do local __t = tostring(check()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
