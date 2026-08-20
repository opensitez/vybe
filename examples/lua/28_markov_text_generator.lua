-- App 28: Markov Text Generator (2-gram)

local corpus = "lua is small and lua is fast and lua is embeddable and fast"

local words = {}
for w in corpus:gmatch("[%a]+") do words[#words + 1] = w end

local chain = {}
for i = 1, #words - 2 do
  local key = words[i] .. " " .. words[i + 1]
  chain[key] = chain[key] or {}
  chain[key][#chain[key] + 1] = words[i + 2]
end

local function pick(t)
  return t[math.random(1, #t)]
end

local function generate(seed_a, seed_b, n)
  local out = {seed_a, seed_b}
  for _ = 1, n do
    local key = out[#out - 1] .. " " .. out[#out]
    local nexts = chain[key]
    if not nexts then break end
    out[#out + 1] = pick(nexts)
  end
  return table.concat(out, " ")
end

print(generate("lua", "is", 15))
