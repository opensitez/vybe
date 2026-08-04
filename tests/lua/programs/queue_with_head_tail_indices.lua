-- vybe-test: lua/programs/queue_with_head_tail_indices
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "abc"
local __i = 0

local q = {head=1, tail=0}
local function enqueue(v) q.tail=q.tail+1; q[q.tail]=v end
local function dequeue() local v=q[q.head]; q[q.head]=nil; q.head=q.head+1; return v end
enqueue('a'); enqueue('b'); enqueue('c')
do local __t = tostring(dequeue() .. dequeue() .. dequeue()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
