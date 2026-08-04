-- vybe-test: lua/scoping_exhaustive/scope_exh_loop_escaped
-- origin: languages/lua/tests/lua/test_scoping_exhaustive.rs

local __w1 = "1\t2\t3"
local __i = 0

local fns = {}
for i = 1, 3 do
  fns[i] = function() return i end
end
do local __t = tostring(fns[1]()) .. "\t" .. tostring(fns[2]()) .. "\t" .. tostring(fns[3]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
