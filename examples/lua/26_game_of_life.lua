-- App 26: Conway's Game of Life

local W, H = 12, 8
local grid = {}
for y = 1, H do
  grid[y] = {}
  for x = 1, W do grid[y][x] = 0 end
end

local seed = {{2,2},{3,3},{1,4},{2,4},{3,4}}
for _, c in ipairs(seed) do grid[c[2]][c[1]] = 1 end

local function count_neighbors(g, x, y)
  local n = 0
  for dy = -1, 1 do
    for dx = -1, 1 do
      if not (dx == 0 and dy == 0) then
        local xx, yy = x + dx, y + dy
        if xx >= 1 and xx <= W and yy >= 1 and yy <= H and g[yy][xx] == 1 then
          n = n + 1
        end
      end
    end
  end
  return n
end

local function step(g)
  local out = {}
  for y = 1, H do
    out[y] = {}
    for x = 1, W do
      local alive = g[y][x] == 1
      local n = count_neighbors(g, x, y)
      out[y][x] = ((alive and (n == 2 or n == 3)) or ((not alive) and n == 3)) and 1 or 0
    end
  end
  return out
end

local function draw(g, gen)
  print("generation", gen)
  for y = 1, H do
    local row = {}
    for x = 1, W do row[#row + 1] = g[y][x] == 1 and "#" or "." end
    print(table.concat(row))
  end
end

for gen = 1, 6 do
  draw(grid, gen)
  grid = step(grid)
end
