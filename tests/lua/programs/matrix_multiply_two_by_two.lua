-- vybe-test: lua/programs/matrix_multiply_two_by_two
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "19,22,43,50"
local __i = 0

local function mmul(a, b)
  return {
    {a[1][1]*b[1][1]+a[1][2]*b[2][1], a[1][1]*b[1][2]+a[1][2]*b[2][2]},
    {a[2][1]*b[1][1]+a[2][2]*b[2][1], a[2][1]*b[1][2]+a[2][2]*b[2][2]},
  }
end
local A = {{1,2},{3,4}}
local B = {{5,6},{7,8}}
local C = mmul(A, B)
do local __t = tostring(C[1][1] .. ',' .. C[1][2] .. ',' .. C[2][1] .. ',' .. C[2][2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
