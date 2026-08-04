-- vybe-test: lua/strings/concat_number_and_string_for_log_line
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "value=42"
local __i = 0

do local __t = tostring("value=" .. 42); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
