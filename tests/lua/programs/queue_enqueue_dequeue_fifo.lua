-- vybe-test: lua/programs/queue_enqueue_dequeue_fifo
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "3"
local __i = 0

local q = {}
local head, tail = 1, 0
local function enqueue(v) tail = tail + 1 q[tail] = v end
local function dequeue() local v = q[head] head = head + 1 return v end
enqueue(1) enqueue(2)
do local __t = tostring(dequeue() + dequeue()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
