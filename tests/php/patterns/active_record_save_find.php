<?php
// vybe-test: php/patterns/active_record_save_find
// origin: languages/php/tests/php/test_patterns.rs
// vybe-test-mode: compile

class Model {
    protected static $table = 'models';
    protected static $records = [];
    protected $attributes = [];
    public function __construct(array $attrs = []) { $this->attributes = $attrs; }
    public function __get($key) { return $this->attributes[$key] ?? null; }
    public function __set($key, $val) { $this->attributes[$key] = $val; }
    public function save(): void {
        $id = $this->attributes['id'] ?? count(static::$records) + 1;
        $this->attributes['id'] = $id;
        static::$records[$id] = $this;
    }
    public static function find(int $id): ?static {
        return static::$records[$id] ?? null;
    }
}
class Post extends Model {
    protected static $table = 'posts';
    protected static $records = [];
}
$p = new Post(['title' => 'Hello', 'body' => 'World']);
$p->save();
echo Post::find(1)->title;
