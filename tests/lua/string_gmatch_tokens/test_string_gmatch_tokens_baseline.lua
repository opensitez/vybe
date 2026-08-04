-- vybe-test: lua/string_gmatch_tokens/test_string_gmatch_tokens_baseline
-- origin: languages/lua/tests/lua/test_string_gmatch_tokens.rs

local __w1 = "true"
local __i = 0

local c = 0
for _ in string.gmatch("w0 w1", "[%a]+") do c = c + 1 end
do local __t = tostring(c == 2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
