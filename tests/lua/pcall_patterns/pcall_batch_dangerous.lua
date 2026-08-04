-- vybe-test: lua/pcall_patterns/pcall_batch_dangerous
-- origin: languages/lua/tests/lua/test_pcall_patterns.rs

local __w1 = "6,err,14"
local __i = 0

local function dangerous(n)
  if n > 10 then error("too big") end
  return n * 2
end
local results = {}
for _, v in ipairs({3, 15, 7}) do
  local ok, r = pcall(dangerous, v)
  results[#results+1] = ok and r or "err"
end
do local __t = tostring(table.concat(results, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
