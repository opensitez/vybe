-- vybe-test: lua/programs/fizzbuzz_one_line_for_fifteen
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "fizzbuzz"
local __i = 0

local n = 15
if n % 15 == 0 then do local __t = tostring("fizzbuzz"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
elseif n % 3 == 0 then do local __t = tostring("fizz"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
elseif n % 5 == 0 then do local __t = tostring("buzz"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
else do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
