-- vybe-test: lua/loops_for_generic/test_for_gen_closure_capture
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "ab"
local __i = 0

local s=''; for k,v in ipairs({'a','b'}) do s=s..v end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
