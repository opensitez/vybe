// Database demo — SQLite from JavaScript
let conn = db.connect("sqlite:///tmp/vybe_demo.db?mode=rwc");

if (conn !== null) {
    // Create table
    db.execute(conn, "CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, name TEXT, age INTEGER)");

    // Insert data
    db.execute(conn, "DELETE FROM users");
    db.execute(conn, "INSERT INTO users (name, age) VALUES (?, ?)", ["Alice", 30]);
    db.execute(conn, "INSERT INTO users (name, age) VALUES (?, ?)", ["Bob", 25]);
    db.execute(conn, "INSERT INTO users (name, age) VALUES (?, ?)", ["Charlie", 35]);

    // Query
    let users = db.query(conn, "SELECT name, age FROM users ORDER BY name");
    console.log(`Found ${users.length} users:`);
    users.forEach((user) => {
        console.log(`  ${user.name}, age ${user.age}`);
    });

    // Scalar
    let count = db.scalar(conn, "SELECT COUNT(*) FROM users");
    console.log(`Total: ${count}`);

    // Tables
    let tables = db.tables(conn);
    console.log(`Tables: ${tables.join(", ")}`);

    // Columns
    let cols = db.columns(conn, "users");
    console.log(`Columns: ${cols.join(", ")}`);

    db.close(conn);
    console.log("Done!");
} else {
    console.log("Failed to connect");
}
