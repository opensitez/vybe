-- vybe-test: lua/programs/stack_implemented_with_table_push_pop
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "3,2,1"
local __i = 0

local stack = {}
local function push(v) stack[#stack+1] = v end
local function pop() local v = stack[#stack]; stack[#stack] = nil; return v end
push(1); push(2); push(3)
do local __t = tostring(pop() .. ',' .. pop() .. ',' .. pop()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
