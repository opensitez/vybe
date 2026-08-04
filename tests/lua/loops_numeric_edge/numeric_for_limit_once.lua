-- vybe-test: lua/loops_numeric_edge/numeric_for_limit_once
-- origin: languages/lua/tests/lua/test_loops_numeric_edge.rs

local __w1 = "1"
local __i = 0

local calls = 0
local function limit() calls = calls + 1; return 3 end
for i = 1, limit() do end
do local __t = tostring(calls); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
