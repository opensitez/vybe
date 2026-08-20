-- App 33: Recommendation Engine (user-based collaborative filtering)

local ratings = {
  alice = {book = 5, game = 1, movie = 4},
  bob = {book = 4, game = 2, movie = 5, course = 5},
  clara = {book = 1, game = 5, movie = 2, course = 1},
}

local function similarity(a, b)
  local dot, na, nb = 0, 0, 0
  for item, ra in pairs(a) do
    local rb = b[item]
    if rb then dot = dot + ra * rb end
    na = na + ra * ra
  end
  for _, rb in pairs(b) do nb = nb + rb * rb end
  if na == 0 or nb == 0 then return 0 end
  return dot / (math.sqrt(na) * math.sqrt(nb))
end

local target = "alice"
local scores = {}

for user, prefs in pairs(ratings) do
  if user ~= target then
    local sim = similarity(ratings[target], prefs)
    for item, r in pairs(prefs) do
      if ratings[target][item] == nil then
        scores[item] = (scores[item] or 0) + sim * r
      end
    end
  end
end

for item, s in pairs(scores) do
  print("recommend", item, string.format("%.3f", s))
end
