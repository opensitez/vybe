-- vybe-test: lua/programs/run_length_encoding_counts
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "a3b2c1"
local __i = 0

local s = "aaabbc"
local out = ""
local i = 1
while i <= #s do
  local c = string.sub(s, i, i)
  local n = 1
  while string.sub(s, i + n, i + n) == c do n = n + 1 end
  out = out .. c .. n
  i = i + n
end
do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
