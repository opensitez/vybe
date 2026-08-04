-- vybe-test: lua/loops_for_generic/test_for_gen_local_scope
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "99 88"
local __i = 0

local k=99; local v=88; for k,v in ipairs({1}) do end; do local __t = tostring(k..' '..v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
