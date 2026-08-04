-- vybe-test: lua/operators_length/test_len_table_hash
-- origin: languages/lua/tests/lua/test_operators_length.rs

local __w1 = "0"
local __i = 0

local t={a=1, b=2}; do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
