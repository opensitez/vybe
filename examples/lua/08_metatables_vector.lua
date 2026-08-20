-- App 08: 2D Physics Sandbox (vector type + collision checks)

local Vec = {}
Vec.__index = Vec

function Vec.new(x, y) return setmetatable({x = x, y = y}, Vec) end
function Vec:__add(o) return Vec.new(self.x + o.x, self.y + o.y) end
function Vec:__sub(o) return Vec.new(self.x - o.x, self.y - o.y) end
function Vec:__mul(k) return Vec.new(self.x * k, self.y * k) end
function Vec:__len() return math.sqrt(self.x * self.x + self.y * self.y) end
function Vec:__tostring() return string.format("(%.2f, %.2f)", self.x, self.y) end

local body = {
  pos = Vec.new(0, 5),
  vel = Vec.new(2, 0),
  acc = Vec.new(0, -9.8),
  radius = 0.5,
}

local dt = 0.1
for tick = 1, 30 do
  body.vel = body.vel + body.acc * dt
  body.pos = body.pos + body.vel * dt

  if body.pos.y - body.radius <= 0 then
    body.pos.y = body.radius
    body.vel.y = -body.vel.y * 0.7
  end

  print(string.format("t=%02d pos=%s speed=%.2f", tick, tostring(body.pos), #body.vel))
end
