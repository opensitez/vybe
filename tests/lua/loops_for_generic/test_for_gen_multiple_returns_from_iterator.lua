-- vybe-test: lua/loops_for_generic/test_for_gen_multiple_returns_from_iterator
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "1AB2AB"
local __i = 0

local function it() local i=0; return function() i=i+1; if i<3 then return i, 'A', 'B' end end end;
         local s=''; for a,b,c in it() do s=s..a..b..c end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
