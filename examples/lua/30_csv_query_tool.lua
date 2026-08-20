-- App 30: CSV Query Tool

local csv = [[
id,name,team,score
1,Ada,red,88
2,Ben,blue,72
3,Cia,red,95
4,Dax,blue,81
]]

local function parse_csv(text)
  local rows = {}
  local headers
  for line in text:gmatch("[^\n]+") do
    local cols = {}
    for cell in line:gmatch("[^,]+") do cols[#cols + 1] = cell end
    if not headers then
      headers = cols
    else
      local row = {}
      for i, h in ipairs(headers) do row[h] = cols[i] end
      row.score = tonumber(row.score)
      rows[#rows + 1] = row
    end
  end
  return rows
end

local rows = parse_csv(csv)

local filtered = {}
for _, r in ipairs(rows) do
  if r.team == "red" and r.score >= 90 then filtered[#filtered + 1] = r end
end

table.sort(filtered, function(a, b) return a.score > b.score end)
for _, r in ipairs(filtered) do
  print(string.format("%s (%s) score=%d", r.name, r.team, r.score))
end
