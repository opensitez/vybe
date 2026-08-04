-- vybe-test: lua/programs/palindrome_string_check
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "true"
local __i = 0

local s = "aba"
local ok = true
for i = 1, #s do
  if string.sub(s, i, i) ~= string.sub(s, #s - i + 1, #s - i + 1) then ok = false break end
end
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
