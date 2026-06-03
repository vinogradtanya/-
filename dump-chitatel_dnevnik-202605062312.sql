
DROP TABLE IF EXISTS `book_goals`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `book_goals` (
  `book_id` int NOT NULL,
  `target_days` int NOT NULL DEFAULT '0',
  `deadline_date` date DEFAULT NULL,
  PRIMARY KEY (`book_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


DROP TABLE IF EXISTS `book_tags`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `book_tags` (
  `book_id` int NOT NULL,
  `tag` varchar(80) NOT NULL,
  PRIMARY KEY (`book_id`,`tag`),
  KEY `idx_book_tags_tag` (`tag`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;
/*!40101 SET character_set_client = @saved_cs_client */;


DROP TABLE IF EXISTS `books`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `books` (
  `id` int NOT NULL AUTO_INCREMENT,
  `title` varchar(500) COLLATE utf8mb4_unicode_ci NOT NULL,
  `author` varchar(300) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '',
  `genre` varchar(200) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '',
  `page_count` int unsigned NOT NULL DEFAULT '0',
  `cover_path` varchar(1024) COLLATE utf8mb4_unicode_ci DEFAULT NULL,
  `date_started` date DEFAULT NULL,
  `date_finished` date DEFAULT NULL,
  `format_type` enum('paper','ebook','audiobook') COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT 'paper',
  `plan_year` smallint unsigned DEFAULT NULL,
  `plan_month` tinyint unsigned DEFAULT NULL,
  `created_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `reading_status` varchar(20) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT 'planned',
  `page_start` int DEFAULT '0',
  `page_end` int DEFAULT '0',
  PRIMARY KEY (`id`),
  KEY `idx_plan` (`plan_year`,`plan_month`)
) ENGINE=InnoDB AUTO_INCREMENT=13 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

--

DROP TABLE IF EXISTS `daily_reading`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `daily_reading` (
  `id` int NOT NULL AUTO_INCREMENT,
  `book_id` int NOT NULL,
  `read_date` date NOT NULL,
  `page_start` int DEFAULT '0',
  `page_end` int DEFAULT '0',
  `pages_read` int unsigned NOT NULL DEFAULT '0',
  `minutes_read` int unsigned NOT NULL DEFAULT '0',
  `note` varchar(500) COLLATE utf8mb4_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `uq_book_date` (`book_id`,`read_date`),
  KEY `idx_date` (`read_date`),
  CONSTRAINT `daily_reading_ibfk_1` FOREIGN KEY (`book_id`) REFERENCES `books` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=49 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;




DROP TABLE IF EXISTS `quotes`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `quotes` (
  `id` int NOT NULL AUTO_INCREMENT,
  `book_id` int NOT NULL,
  `quote_text` text COLLATE utf8mb4_unicode_ci NOT NULL,
  `page_number` int DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_book` (`book_id`),
  CONSTRAINT `quotes_ibfk_1` FOREIGN KEY (`book_id`) REFERENCES `books` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


DROP TABLE IF EXISTS `reading_goals`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `reading_goals` (
  `scope_year` int NOT NULL,
  `scope_month` tinyint NOT NULL DEFAULT '0',
  `target_pages` int NOT NULL DEFAULT '0',
  `target_minutes` int NOT NULL DEFAULT '0',
  PRIMARY KEY (`scope_year`,`scope_month`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


DROP TABLE IF EXISTS `reviews`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `reviews` (
  `id` int NOT NULL AUTO_INCREMENT,
  `book_id` int NOT NULL,
  `rating_idea` tinyint unsigned NOT NULL DEFAULT '0',
  `rating_plot` tinyint unsigned NOT NULL DEFAULT '0',
  `rating_characters` tinyint unsigned NOT NULL DEFAULT '0',
  `rating_author_skill` tinyint unsigned NOT NULL DEFAULT '0',
  `review_text` text COLLATE utf8mb4_unicode_ci,
  PRIMARY KEY (`id`),
  UNIQUE KEY `book_id` (`book_id`),
  CONSTRAINT `reviews_ibfk_1` FOREIGN KEY (`book_id`) REFERENCES `books` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=73 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
