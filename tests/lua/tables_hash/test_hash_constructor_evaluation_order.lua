-- vybe-test: lua/tables_hash/test_hash_constructor_evaluation_order
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "10"
local __i = 0

local i=0; local function f() i=i+1; return i end; local t={[f()]=f(), [f()]=f()}; local cnt=0; for k,v in pairs(t) do cnt=cnt+k+v end; do local __t = tostring(cnt); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
