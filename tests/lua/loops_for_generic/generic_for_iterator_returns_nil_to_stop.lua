-- vybe-test: lua/loops_for_generic/generic_for_iterator_returns_nil_to_stop
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "321"
local __i = 0

local function countdown(from, cur)
  if cur == nil then cur = from end
  if cur <= 0 then return nil end
  return cur - 1, cur
end
local s = ''
for _, v in countdown, 3 do s = s .. v end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
