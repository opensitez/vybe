-- vybe-test: lua/string_patterns_complex/test_pattern_multiple_captures_gmatch
-- origin: languages/lua/tests/lua/test_string_patterns_complex.rs

local __w1 = "key1:value1 key2:value2 "
local __i = 0

local s = 'key1=value1 key2=value2'
local res = ''
for k, v in string.gmatch(s, '(%w+)=(%w+)') do
    res = res .. k .. ':' .. v .. ' '
end
do local __t = tostring(res); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
