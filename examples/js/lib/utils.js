export function capitalize(str) {
    if (str.length === 0) return str;
    return str.charAt(0).toUpperCase() + str.slice(1);
}

export function repeat(str, n) {
    let result = "";
    for (let i = 0; i < n; i++) {
        result = result + str;
    }
    return result;
}

export let VERSION = "1.0.0";
