CREATE TABLE `account_guard` (
	`account_id` MEDIUMINT(7) UNSIGNED NOT NULL,
	`code` CHAR(5) NULL DEFAULT '',
	`timestamp` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
	`last_send_email` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
	PRIMARY KEY (`account_id`) USING BTREE
)
ENGINE=InnoDB
;