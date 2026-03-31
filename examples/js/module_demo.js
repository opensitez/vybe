// Module demo — importing from local files
import { capitalize, repeat, VERSION } from "./lib/utils.js";
import { clamp, lerp, distance } from "./lib/math.js";

console.log(`Utils v${VERSION}`);
console.log(capitalize("hello world"));
console.log(repeat("ab", 3));
console.log(`clamp(15, 0, 10) = ${clamp(15, 0, 10)}`);
console.log(`lerp(0, 100, 0.5) = ${lerp(0, 100, 0.5)}`);
console.log(`distance(0,0,3,4) = ${distance(0, 0, 3, 4)}`);
