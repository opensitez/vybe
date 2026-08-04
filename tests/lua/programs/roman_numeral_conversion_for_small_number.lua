-- vybe-test: lua/programs/roman_numeral_conversion_for_small_number
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "MMXXIV"
local __i = 0

local function to_roman(n)
  local vals = {1000,900,500,400,100,90,50,40,10,9,5,4,1}
  local syms = {'M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'}
  local result = ''
  for i, v in ipairs(vals) do
    while n >= v do result = result .. syms[i]; n = n - v end
  end
  return result
end
do local __t = tostring(to_roman(2024)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
