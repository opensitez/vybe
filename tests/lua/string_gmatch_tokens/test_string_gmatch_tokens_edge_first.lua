-- vybe-test: lua/string_gmatch_tokens/test_string_gmatch_tokens_edge_first
-- origin: languages/lua/tests/lua/test_string_gmatch_tokens.rs

local __w1 = "true"
local __i = 0

local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10 w11 w12 w13 w14 w15 w16", "[%a]+") do c = c + 1 end
do local __t = tostring(c == 17); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
