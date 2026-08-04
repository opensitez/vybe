-- vybe-test: lua/vararg/table_pack_preserves_nil_count
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "3,nil"
local __i = 0

local t = table.pack(10, nil, 30)
do local __t = tostring(t.n .. ',' .. tostring(t[2])); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
