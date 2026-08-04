-- vybe-test: lua/loops_for_generic/test_for_gen_eval_iterator_once
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "1020 1"
local __i = 0

local c=0; local function get_iter() c=c+1; local i=0; return function() i=i+1; if i<=2 then return i, i*10 end end end;
         local s=''; for k,v in get_iter() do s=s..v end; do local __t = tostring(s..' '..c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
