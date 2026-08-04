-- vybe-test: lua/functions/closure_reads_captured_upvalue
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "0"
local __i = 0

function make()
  local n=0
  return function() return n end
end
do local __t = tostring(make()()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
