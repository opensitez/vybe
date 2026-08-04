-- vybe-test: lua/scoping_exhaustive/scope_exh_param_shadow
-- origin: languages/lua/tests/lua/test_scoping_exhaustive.rs

local __w1 = "99\t1"
local __i = 0

local x = 1
local function f(x) return x end
do local __t = tostring(f(99)) .. "\t" .. tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
