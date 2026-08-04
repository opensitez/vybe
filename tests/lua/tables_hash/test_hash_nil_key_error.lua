-- vybe-test: lua/tables_hash/test_hash_nil_key_error
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "false"
local __i = 0

local t={}; local ok = pcall(function() t[nil]=1 end); do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
