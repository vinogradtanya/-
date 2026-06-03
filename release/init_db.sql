-- Инициализация базы данных для Читательского дневника
CREATE DATABASE IF NOT EXISTS chitatel_dnevnik CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE chitatel_dnevnik;

CREATE TABLE IF NOT EXISTS books (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    title VARCHAR(500) NOT NULL DEFAULT '',
    author VARCHAR(300) NOT NULL DEFAULT '',
    genre VARCHAR(150) NOT NULL DEFAULT '',
    page_count INT NOT NULL DEFAULT 0,
    page_start INT NOT NULL DEFAULT 0,
    page_end INT NOT NULL DEFAULT 0,
    format_type VARCHAR(20) NOT NULL DEFAULT 'paper',
    plan_year INT NULL,
    plan_month TINYINT NULL,
    cover_path VARCHAR(500) NULL,
    date_started DATE NULL,
    date_finished DATE NULL,
    reading_status VARCHAR(20) NOT NULL DEFAULT 'planned',
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS quotes (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    book_id INT NOT NULL,
    quote_text TEXT NOT NULL,
    FOREIGN KEY (book_id) REFERENCES books(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS reviews (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    book_id INT NOT NULL UNIQUE,
    rating_idea TINYINT NOT NULL DEFAULT 3,
    rating_plot TINYINT NOT NULL DEFAULT 3,
    rating_characters TINYINT NOT NULL DEFAULT 3,
    rating_author_skill TINYINT NOT NULL DEFAULT 3,
    review_text TEXT NULL,
    FOREIGN KEY (book_id) REFERENCES books(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS daily_reading (
    book_id INT NOT NULL,
    read_date DATE NOT NULL,
    page_start INT NOT NULL DEFAULT 0,
    page_end INT NOT NULL DEFAULT 0,
    pages_read INT NOT NULL DEFAULT 0,
    minutes_read INT NOT NULL DEFAULT 0,
    note TEXT NULL,
    PRIMARY KEY (book_id, read_date),
    FOREIGN KEY (book_id) REFERENCES books(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS book_tags (
    book_id INT NOT NULL,
    tag VARCHAR(80) NOT NULL,
    PRIMARY KEY (book_id, tag),
    INDEX idx_book_tags_tag (tag),
    FOREIGN KEY (book_id) REFERENCES books(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS reading_goals (
    scope_year INT NOT NULL,
    scope_month TINYINT NOT NULL DEFAULT 0,
    target_pages INT NOT NULL DEFAULT 0,
    target_minutes INT NOT NULL DEFAULT 0,
    PRIMARY KEY (scope_year, scope_month)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS book_goals (
    book_id INT NOT NULL PRIMARY KEY,
    target_days INT NOT NULL DEFAULT 0,
    deadline_date DATE NULL,
    FOREIGN KEY (book_id) REFERENCES books(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
