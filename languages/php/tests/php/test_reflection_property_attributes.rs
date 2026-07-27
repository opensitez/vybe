crate::php_cases! {
    reflection_property_get_attributes => {
        r#"<?php
#[Attribute]
class Column {
    public function __construct(public string $type) {}
}

class User {
    #[Column('integer')]
    public int $id;
    
    #[Column('string')]
    public string $name;
}

$rp1 = new ReflectionProperty(User::class, 'id');
$rp2 = new ReflectionProperty(User::class, 'name');

echo $rp1->getAttributes()[0]->getArguments()[0] . "|";
echo $rp2->getAttributes()[0]->getArguments()[0];
"#,
        ["integer|string"]
    };
}
