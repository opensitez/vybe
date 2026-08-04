-- vybe-test: lua/programs/sum_first_n_integers
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "15"
local __i = 0

local n = 5
local sum = 0
local i = 1
while i <= n do
  sum = sum + i
  i = i + 1
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
