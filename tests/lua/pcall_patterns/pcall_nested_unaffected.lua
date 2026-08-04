-- vybe-test: lua/pcall_patterns/pcall_nested_unaffected
-- origin: languages/lua/tests/lua/test_pcall_patterns.rs

local __w1 = "true"
local __i = 0

local outer_ok = pcall(function()
  local inner_ok = pcall(function() error("inner") end)
  assert(inner_ok == false)
end)
do local __t = tostring(outer_ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
