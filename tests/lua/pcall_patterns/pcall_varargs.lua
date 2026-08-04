-- vybe-test: lua/pcall_patterns/pcall_varargs
-- origin: languages/lua/tests/lua/test_pcall_patterns.rs

local __w1 = "true\t10"
local __i = 0

local function sum(...)
  local s = 0
  for _, v in ipairs({...}) do s = s + v end
  return s
end
local ok, v = pcall(sum, 1, 2, 3, 4)
do local __t = tostring(ok) .. "\t" .. tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
