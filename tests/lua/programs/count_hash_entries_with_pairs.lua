-- vybe-test: lua/programs/count_hash_entries_with_pairs
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "3"
local __i = 0

local cfg = {host = "x", port = 80, ssl = true}
local n = 0
for _ in pairs(cfg) do n = n + 1 end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
