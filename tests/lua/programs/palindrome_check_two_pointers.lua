-- vybe-test: lua/programs/palindrome_check_two_pointers
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "true"
local __i = 0

local s = "radar"
local i, j, ok = 1, #s, true
while i < j do
  if string.sub(s, i, i) ~= string.sub(s, j, j) then ok = false break end
  i, j = i + 1, j - 1
end
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
