-- vybe-test: lua/loops_for_numeric/numeric_for_with_pcall_protection_on_error_in_body
-- origin: languages/lua/tests/lua/test_loops_for_numeric.rs

local __w1 = "false"
local __i = 0

local ok, err = pcall(function()
  for i = 1, 5 do
    if i == 3 then error('stop at 3') end
  end
end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
