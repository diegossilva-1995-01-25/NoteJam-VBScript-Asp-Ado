CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    email VARCHAR(75) NOT NULL,
    password VARCHAR(128) NOT NULL
);

CREATE TABLE IF NOT EXISTS pads (
    id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    name VARCHAR(100) NOT NULL,
    user_id INTEGER NOT NULL REFERENCES users(id)
);

CREATE TABLE IF NOT EXISTS notes (
    id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    pad_id INTEGER REFERENCES pads(id),
    user_id INTEGER NOT NULL REFERENCES users(id),
    name VARCHAR(100) NOT NULL,
    text text NOT NULL,
    created_at DATETIME NOT NULL,
    updated_at DATETIME NOT NULL
);

DROP TABLE IF EXISTS users;
DROP TABLE IF EXISTS pads;
DROP TABLE IF EXISTS notes;

-- Funções CRUD de cada tabela
-- Usuários
INSERT INTO users(email, password) VALUES (?, ?);
	
UPDATE users SET email = ?, password = ? WHERE id = ?;
	
DELETE FROM users WHERE id = ?;

SELECT * FROM users WHERE id = ?;


-- Pastas
INSERT INTO pads(name, user_id) VALUES (?, ?);
	
UPDATE pads SET name = ?, user_id = ? WHERE id = ?;
	
DELETE FROM pads WHERE id = ?;

SELECT name FROM pads INNER JOIN users ON pads.user_id = users.id WHERE id = ?;


-- Notas
INSERT INTO notes(pad_id, user_id, name, text, created_at, updated_at) 
	VALUES (?, ?, ?, ?, DATETIME('NOW', 'LOCALTIME'), DATETIME('NOW', 'LOCALTIME'));
	
UPDATE notes SET pad_id = ?, user_id = ?, name = ?, text = ?, updated_at = DATETIME('now', 'localtime') WHERE id = ?;
	
DELETE FROM notes WHERE id = ?;

SELECT * FROM notes;

SELECT id, name FROM notes INNER JOIN pads ON notes.pad_id = pads.id
	INNER JOIN users ON pads.user_id = users.id WHERE id = ?;

