-- vybe-test: lua/iteration/generic_for_triple_return_from_iterator
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "1x1"
local __i = 0

local function iter(s, k)
  k = (k or 0) + 1
  if k > #s then return nil end
  return k, s[k], #s
end
local a, b, c = iter({"x"}, 0)
do local __t = tostring(a .. b .. c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
