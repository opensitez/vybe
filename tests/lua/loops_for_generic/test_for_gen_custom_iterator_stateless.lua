-- vybe-test: lua/loops_for_generic/test_for_gen_custom_iterator_stateless
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "110220330"
local __i = 0

local function iter(state, var) var=var+1; if var<=state then return var, var*10 end end;
         local s=''; for k,v in iter, 3, 0 do s=s..k..v end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
