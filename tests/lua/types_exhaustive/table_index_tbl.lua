-- vybe-test: lua/types_exhaustive/table_index_tbl
-- origin: languages/lua/tests/lua/test_types_exhaustive.rs

local __w1 = "d"
local __i = 0

local t = {}; local k = {}; t[k] = "d"; do local __t = tostring(t[k]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
