-- vybe-test: lua/closures/closure_passed_as_callback_to_iterating_helper
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "20"
local __i = 0

local function map(t, f)
  local out = {}
  for i, v in ipairs(t) do out[i] = f(v) end
  return out
end
do local __t = tostring(map({1,2}, function(x) return x * 10 end)[2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
