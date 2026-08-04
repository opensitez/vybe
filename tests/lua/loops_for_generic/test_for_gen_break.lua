-- vybe-test: lua/loops_for_generic/test_for_gen_break
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "1020"
local __i = 0

local s=''; for k,v in ipairs({10,20,30,40}) do s=s..v; if k==2 then break end end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
