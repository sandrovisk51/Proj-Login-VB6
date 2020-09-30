/*
Navicat MySQL Data Transfer

Source Server         : MYSQL5.1
Source Server Version : 50172
Source Host           : localhost:3306
Source Database       : dblogin

Target Server Type    : MYSQL
Target Server Version : 50172
File Encoding         : 65001

Date: 2020-09-29 19:31:32
*/

SET FOREIGN_KEY_CHECKS=0;

-- ----------------------------
-- Table structure for `tblogin`
-- ----------------------------
DROP TABLE IF EXISTS `tblogin`;
CREATE TABLE `tblogin` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `email` varchar(250) DEFAULT NULL,
  `senha` varchar(250) DEFAULT NULL,
  `data` datetime DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of tblogin
-- ----------------------------
INSERT INTO `tblogin` VALUES ('1', 'teste@email.com', '123456', '2020-09-29 11:52:31');
INSERT INTO `tblogin` VALUES ('2', 'kiko@email.com', '1234', '2020-09-29 11:53:18');
