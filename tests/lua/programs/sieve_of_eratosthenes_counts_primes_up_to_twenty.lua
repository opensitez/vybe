-- vybe-test: lua/programs/sieve_of_eratosthenes_counts_primes_up_to_twenty
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "8"
local __i = 0

local n = 20
local is_prime = {}
for i = 2, n do is_prime[i] = true end
for p = 2, math.floor(math.sqrt(n)) do
  if is_prime[p] then
    for m = p * p, n, p do is_prime[m] = false end
  end
end
local count = 0
for i = 2, n do if is_prime[i] then count = count + 1 end end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
