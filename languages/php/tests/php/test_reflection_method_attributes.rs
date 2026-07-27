crate::php_cases! {
    reflection_method_get_attributes => {
        r#"<?php
#[Attribute]
class Route {
    public function __construct(public string $path) {}
}

class Controller {
    #[Route('/home')]
    public function index() {}
}

$rm = new ReflectionMethod(Controller::class, 'index');
$attrs = $rm->getAttributes();
echo $attrs[0]->getName() . "->";
echo $attrs[0]->getArguments()[0];
"#,
        ["Route->/home"]
    };

    reflection_method_get_attributes_filtered => {
        r#"<?php
#[Attribute] class Get {}
#[Attribute] class Post {}

class Api {
    #[Get]
    #[Post]
    public function endpoint() {}
}

$rm = new ReflectionMethod(Api::class, 'endpoint');
$getAttrs = $rm->getAttributes(Get::class);
echo count($getAttrs) . "|";
$postAttrs = $rm->getAttributes(Post::class);
echo count($postAttrs) . "|";
$missingAttrs = $rm->getAttributes('Missing');
echo count($missingAttrs);
"#,
        ["1|1|0"]
    };
}
