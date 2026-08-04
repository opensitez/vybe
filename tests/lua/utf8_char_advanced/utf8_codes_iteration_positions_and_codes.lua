-- vybe-test: lua/utf8_char_advanced/utf8_codes_iteration_positions_and_codes
-- origin: languages/lua/tests/lua/test_utf8_char_advanced.rs

local __w1 = "1:945 3:946 "
local __i = 0

local s = "αβ"
local r = ""
for p, c in utf8.codes(s) do r = r .. p .. ":" .. c .. " " end
do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
