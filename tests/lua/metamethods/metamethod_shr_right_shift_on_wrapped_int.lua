-- vybe-test: lua/metamethods/metamethod_shr_right_shift_on_wrapped_int
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "4"
local __i = 0

local mt = {__shr = function(a, b) return {n = a.n >> b.n} end}
do local __t = tostring((setmetatable({n=8}, mt) >> setmetatable({n=1}, mt)).n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
