-- vybe-test: lua/programs/balanced_parentheses_checker
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "true"
local __i = 0

local s = "(()())"
local depth = 0
local ok = true
for i = 1, #s do
  local c = string.sub(s, i, i)
  if c == "(" then depth = depth + 1 elseif c == ")" then depth = depth - 1 end
  if depth < 0 then ok = false break end
end
do local __t = tostring(ok and depth == 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
