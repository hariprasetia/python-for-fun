SET NAMES utf8mb4;
SET FOREIGN_KEY_CHECKS = 0;

-- ----------------------------
-- Table structure for users
-- ----------------------------
DROP TABLE IF EXISTS `users`;
CREATE TABLE `users` (
  `userid` int NOT NULL,
  `nama` char(50) DEFAULT NULL,
  `tinggi_badan` double DEFAULT NULL,
  `berat_badan` double DEFAULT NULL,
  `tanggal_lahir` date DEFAULT NULL,
  `golongan_darah` char(10) DEFAULT NULL,
  `alamat` varchar(255) DEFAULT NULL,
  `telepon` char(20) DEFAULT NULL,
  PRIMARY KEY (`userid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

SET FOREIGN_KEY_CHECKS = 1;
