-- vybe-test: lua/control_flow/return_value_from_inside_if_branch
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "negative,zero,positive"
local __i = 0

local function classify(n)
  if n < 0 then return 'negative'
  elseif n == 0 then return 'zero'
  else return 'positive'
  end
end
do local __t = tostring(classify(-5) .. ',' .. classify(0) .. ',' .. classify(3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
