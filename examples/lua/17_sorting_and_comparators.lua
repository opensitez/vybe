-- 17_sorting_and_comparators.lua
-- Demonstrates custom sort comparator, stable-style tie-breaking, and grouped reporting.

local players = {
  {name = "Mira", score = 1200, level = 8},
  {name = "Ari", score = 1200, level = 9},
  {name = "Tao", score = 980, level = 11},
  {name = "Bea", score = 1600, level = 6}
}

table.sort(players, function(a, b)
  if a.score ~= b.score then
    return a.score > b.score
  end
  if a.level ~= b.level then
    return a.level > b.level
  end
  return a.name < b.name
end)

for rank, p in ipairs(players) do
  print(string.format("#%d %-4s score=%d level=%d", rank, p.name, p.score, p.level))
end
