-- vybe-test: lua/functions/function_return_drives_while_condition
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "0"
local __i = 0

function has_more(n) return n > 0 end
local n = 2
while has_more(n) do n = n - 1 end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
