-- vybe-test: lua/utf8/utf8_codes_yields_position_and_codepoint
-- origin: languages/lua/tests/lua/test_utf8.rs

local __w1 = "1:97"
local __i = 0

for p, c in utf8.codes("a") do do local __t = tostring(p .. ":" .. c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
