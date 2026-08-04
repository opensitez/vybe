-- vybe-test: lua/functions/closure_mutates_enclosed_local_across_calls
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "1,2"
local __i = 0

function counter()
  local n=0
  return function() n=n+1 return n end
end
local c=counter()
do local __t = tostring(c() .. "," .. c()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
