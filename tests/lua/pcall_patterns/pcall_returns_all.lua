-- vybe-test: lua/pcall_patterns/pcall_returns_all
-- origin: languages/lua/tests/lua/test_pcall_patterns.rs

local __w1 = "true\t1\t2\t3"
local __i = 0

local function multi() return 1, 2, 3 end
local ok, a, b, c = pcall(multi)
do local __t = tostring(ok) .. "\t" .. tostring(a) .. "\t" .. tostring(b) .. "\t" .. tostring(c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
