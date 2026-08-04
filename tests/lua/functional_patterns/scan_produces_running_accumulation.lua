-- vybe-test: lua/functional_patterns/scan_produces_running_accumulation
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "0,1,3,6,10"
local __i = 0

local function scan(t, f, init)
  local r = {init}
  for _, v in ipairs(t) do
    r[#r+1] = f(r[#r], v)
  end
  return r
end
local s = scan({1,2,3,4}, function(a, b) return a + b end, 0)
do local __t = tostring(table.concat(s, ',')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
