-- phpMyAdmin SQL Dump
-- version 4.4.15.10
-- https://www.phpmyadmin.net
--
-- Servidor: localhost
-- Tiempo de generación: 26-04-2021 a las 21:46:02
-- Versión del servidor: 5.5.68-MariaDB
-- Versión de PHP: 5.4.16

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `walter`
--

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `clanes`
--

CREATE TABLE IF NOT EXISTS `clanes` (
  `Id` int(11) NOT NULL,
  `Founder` varchar(20) DEFAULT NULL,
  `GuildName` varchar(20) DEFAULT NULL,
  `Fecha` date DEFAULT NULL,
  `Antifaccion` bigint(20) DEFAULT '0',
  `Alineacion` tinyint(4) DEFAULT '0',
  `Desc` varchar(500) DEFAULT NULL,
  `GuildNews` varchar(500) DEFAULT NULL,
  `URL` varchar(80) DEFAULT NULL,
  `Leader` varchar(20) DEFAULT NULL,
  `Codex1` varchar(80) DEFAULT NULL,
  `Codex2` varchar(80) DEFAULT NULL,
  `Codex3` varchar(80) DEFAULT NULL,
  `Codex4` varchar(80) DEFAULT NULL,
  `Codex5` varchar(80) DEFAULT NULL,
  `Codex6` varchar(80) DEFAULT NULL,
  `Codex7` varchar(80) DEFAULT NULL,
  `Codex8` varchar(80) DEFAULT NULL,
  `CantMiembros` int(11) DEFAULT '0',
  `EleccionesFinalizan` date DEFAULT NULL
) ENGINE=MyISAM AUTO_INCREMENT=13 DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `clanes_propuestas`
--

CREATE TABLE IF NOT EXISTS `clanes_propuestas` (
  `Id` bigint(20) NOT NULL,
  `IdClan` int(11) DEFAULT NULL,
  `IdClanTo` int(11) DEFAULT NULL,
  `Detalle` varchar(400) DEFAULT NULL,
  `Tipo` int(11) DEFAULT '0'
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `clanes_relaciones`
--

CREATE TABLE IF NOT EXISTS `clanes_relaciones` (
  `IdClan` int(11) NOT NULL,
  `IdClanTo` int(11) NOT NULL,
  `Relacion` int(11) DEFAULT '0'
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `clanes_solicitudes`
--

CREATE TABLE IF NOT EXISTS `clanes_solicitudes` (
  `Id` bigint(20) NOT NULL,
  `IdClan` int(11) DEFAULT NULL,
  `Nombre` varchar(20) DEFAULT NULL,
  `Solicitud` varchar(400) DEFAULT NULL
) ENGINE=MyISAM AUTO_INCREMENT=159 DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `clanes_votos`
--

CREATE TABLE IF NOT EXISTS `clanes_votos` (
  `IdClan` int(11) NOT NULL,
  `Nombre` varchar(20) NOT NULL,
  `Voto` varchar(20) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `conexiones`
--

CREATE TABLE IF NOT EXISTS `conexiones` (
  `Id` bigint(20) NOT NULL,
  `IdPj` bigint(20) NOT NULL,
  `IP` varchar(15) DEFAULT NULL,
  `Fecha` datetime DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `cuentas`
--

CREATE TABLE IF NOT EXISTS `cuentas` (
  `Id` bigint(20) NOT NULL,
  `Nombre` varchar(20) NOT NULL,
  `Password` varchar(32) DEFAULT NULL,
  `Email` varchar(100) DEFAULT NULL,
  `NombreApellido` varchar(150) DEFAULT NULL,
  `Direccion` varchar(250) DEFAULT NULL,
  `Ciudad` varchar(40) DEFAULT NULL,
  `Pais` varchar(2) DEFAULT NULL,
  `Telefono` varchar(20) DEFAULT NULL,
  `Pregunta` varchar(100) DEFAULT NULL,
  `Respuesta` varchar(50) DEFAULT NULL,
  `EmailAux` varchar(100) DEFAULT NULL,
  `Ban` tinyint(1) DEFAULT '0',
  `Nacimiento` date DEFAULT NULL
) ENGINE=MyISAM AUTO_INCREMENT=1329 DEFAULT CHARSET=utf8;

--
-- Volcado de datos para la tabla `cuentas`
--

INSERT INTO `cuentas` (`Id`, `Nombre`, `Password`, `Email`, `NombreApellido`, `Direccion`, `Ciudad`, `Pais`, `Telefono`, `Pregunta`, `Respuesta`, `EmailAux`, `Ban`, `Nacimiento`) VALUES
(1325, 'deidhara', '81dc9bdb52d04dc20036dbd8313ed055', 'as@as.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1326, 'natalia', '81dc9bdb52d04dc20036dbd8313ed055', 'nata@ass.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1327, 'pochola', '81dc9bdb52d04dc20036dbd8313ed055', 'pochola@pochola.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1328, 'asd', '7815696ecbf1c96e6894b779456d330e', 'asas@a.sas', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `fortalezas`
--

CREATE TABLE IF NOT EXISTS `fortalezas` (
  `Id` int(11) NOT NULL,
  `Nombre` varchar(50) DEFAULT NULL,
  `IdClan` int(11) DEFAULT NULL,
  `Fecha` datetime DEFAULT NULL,
  `X` int(11) DEFAULT NULL,
  `Y` int(11) DEFAULT NULL,
  `SpawnX` int(11) DEFAULT '0',
  `SpawnY` int(11) DEFAULT NULL,
  `NPCRey` int(11) DEFAULT NULL
) ENGINE=MyISAM AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

--
-- Volcado de datos para la tabla `fortalezas`
--

INSERT INTO `fortalezas` (`Id`, `Nombre`, `IdClan`, `Fecha`, `X`, `Y`, `SpawnX`, `SpawnY`, `NPCRey`) VALUES
(1, 'Oeste', 11, '2018-12-15 14:50:15', 40, 1460, 18, 1304, 620),
(2, 'Este', 11, '2018-12-15 14:47:58', 1055, 1460, 1080, 1252, 619);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `penas`
--

CREATE TABLE IF NOT EXISTS `penas` (
  `Id` tinyint(20) NOT NULL,
  `IdPj` bigint(20) DEFAULT NULL,
  `Razon` varchar(100) DEFAULT NULL,
  `Fecha` datetime NOT NULL,
  `IdGM` bigint(20) DEFAULT NULL,
  `Tiempo` int(11) DEFAULT '0'
) ENGINE=MyISAM AUTO_INCREMENT=128 DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `pjs`
--

CREATE TABLE IF NOT EXISTS `pjs` (
  `Id` bigint(20) unsigned NOT NULL,
  `IdAccount` bigint(20) DEFAULT '0',
  `Nombre` varchar(20) DEFAULT NULL,
  `Genero` tinyint(4) DEFAULT NULL,
  `Raza` tinyint(4) DEFAULT NULL,
  `Hogar` tinyint(4) DEFAULT NULL,
  `Clase` tinyint(4) DEFAULT NULL,
  `Heading` tinyint(4) DEFAULT NULL,
  `Head` int(11) DEFAULT NULL,
  `Body` int(11) DEFAULT NULL,
  `Arma` int(11) DEFAULT NULL,
  `Escudo` int(11) DEFAULT NULL,
  `Casco` int(11) DEFAULT NULL,
  `Uptime` bigint(20) DEFAULT NULL,
  `LastIP` varchar(15) DEFAULT NULL,
  `LastConnect` datetime DEFAULT NULL,
  `Map` tinyint(4) DEFAULT NULL,
  `X` int(11) DEFAULT NULL,
  `Y` int(11) DEFAULT NULL,
  `Logged` tinyint(1) DEFAULT NULL,
  `Muerto` tinyint(1) DEFAULT NULL,
  `Escondido` tinyint(1) DEFAULT NULL,
  `Hambre` tinyint(1) DEFAULT NULL,
  `Sed` tinyint(1) DEFAULT NULL,
  `Desnudo` tinyint(1) DEFAULT NULL,
  `Ban` tinyint(1) DEFAULT NULL,
  `Navegando` tinyint(1) DEFAULT NULL,
  `Envenenado` tinyint(1) DEFAULT NULL,
  `Paralizado` tinyint(1) DEFAULT NULL,
  `PerteneceReal` tinyint(1) DEFAULT NULL,
  `PerteneceCaos` tinyint(1) DEFAULT NULL,
  `Pena` int(11) DEFAULT NULL,
  `EjercitoReal` tinyint(1) DEFAULT NULL,
  `EjercitoCaos` tinyint(1) DEFAULT NULL,
  `CiudMatados` tinyint(1) DEFAULT NULL,
  `CrimMatados` tinyint(1) DEFAULT NULL,
  `rArCaos` tinyint(1) DEFAULT NULL,
  `rArReal` tinyint(1) DEFAULT NULL,
  `rExCaos` tinyint(1) DEFAULT NULL,
  `rExReal` tinyint(1) DEFAULT NULL,
  `recCaos` int(11) DEFAULT NULL,
  `recReal` int(11) DEFAULT NULL,
  `Reenlistadas` int(11) DEFAULT NULL,
  `NivelIngreso` int(11) DEFAULT NULL,
  `FechaIngreso` date DEFAULT NULL,
  `MatadosIngreso` int(11) DEFAULT NULL,
  `NextRecompensa` int(11) DEFAULT NULL,
  `At1` int(11) DEFAULT NULL,
  `At2` int(11) DEFAULT NULL,
  `At3` int(11) DEFAULT NULL,
  `At4` int(11) DEFAULT NULL,
  `At5` int(11) DEFAULT NULL,
  `Sk1` int(11) DEFAULT NULL,
  `Sk2` int(11) DEFAULT NULL,
  `Sk3` int(11) DEFAULT NULL,
  `Sk4` int(11) DEFAULT NULL,
  `Sk5` int(11) DEFAULT NULL,
  `Sk6` int(11) DEFAULT NULL,
  `Sk7` int(11) DEFAULT NULL,
  `Sk8` int(11) DEFAULT NULL,
  `Sk9` int(11) DEFAULT NULL,
  `Sk10` int(11) DEFAULT NULL,
  `Sk11` int(11) DEFAULT NULL,
  `Sk12` int(11) DEFAULT NULL,
  `Sk13` int(11) DEFAULT NULL,
  `Sk14` int(11) DEFAULT NULL,
  `Sk15` int(11) DEFAULT NULL,
  `Sk16` int(11) DEFAULT NULL,
  `Sk17` int(11) DEFAULT NULL,
  `Sk18` int(11) DEFAULT NULL,
  `Sk19` int(11) DEFAULT NULL,
  `Sk20` int(11) DEFAULT NULL,
  `Email` varchar(100) DEFAULT NULL,
  `Gld` bigint(20) DEFAULT NULL,
  `Banco` bigint(20) DEFAULT NULL,
  `MaxHP` int(11) DEFAULT NULL,
  `MinHP` int(11) DEFAULT NULL,
  `MaxSTA` int(11) DEFAULT NULL,
  `MinSTA` int(11) DEFAULT NULL,
  `MaxMAN` int(11) DEFAULT NULL,
  `MinMAN` int(11) DEFAULT NULL,
  `MaxHIT` int(11) DEFAULT NULL,
  `MinHIT` int(11) DEFAULT NULL,
  `MaxAGU` int(11) DEFAULT NULL,
  `MinAGU` int(11) DEFAULT NULL,
  `MaxHAM` int(11) DEFAULT NULL,
  `MinHAM` int(11) DEFAULT NULL,
  `SkillPtsLibres` int(11) DEFAULT NULL,
  `EXP` bigint(20) DEFAULT NULL,
  `ELV` int(11) DEFAULT NULL,
  `ELU` bigint(20) DEFAULT NULL,
  `UserMuertes` bigint(20) DEFAULT NULL,
  `NpcsMuertes` bigint(20) DEFAULT NULL,
  `WeaponEqpSlot` int(11) DEFAULT NULL,
  `ArmourEqpSlot` int(11) DEFAULT NULL,
  `CascoEqpSlot` int(11) DEFAULT NULL,
  `EscudoEqpSlot` int(11) DEFAULT NULL,
  `BarcoSlot` int(11) DEFAULT NULL,
  `MunicionSlot` int(11) DEFAULT NULL,
  `AnilloSlot` int(11) DEFAULT NULL,
  `Rep_Asesino` bigint(20) DEFAULT NULL,
  `Rep_Bandido` bigint(20) DEFAULT NULL,
  `Rep_Burguesia` bigint(20) DEFAULT NULL,
  `Rep_Ladrones` bigint(20) DEFAULT NULL,
  `Rep_Nobles` bigint(20) DEFAULT NULL,
  `Rep_Plebe` bigint(20) DEFAULT NULL,
  `Rep_Promedio` bigint(20) DEFAULT NULL,
  `NroMascotas` int(11) DEFAULT NULL,
  `Masc1` int(11) DEFAULT NULL,
  `Masc2` int(11) DEFAULT NULL,
  `Masc3` int(11) DEFAULT NULL,
  `TrainningTime` bigint(20) DEFAULT '0',
  `H1` int(11) DEFAULT NULL,
  `H2` int(11) DEFAULT NULL,
  `H3` int(11) DEFAULT NULL,
  `H4` int(11) DEFAULT NULL,
  `H5` int(11) DEFAULT NULL,
  `H6` int(11) DEFAULT NULL,
  `H7` int(11) DEFAULT NULL,
  `H8` int(11) DEFAULT NULL,
  `H9` int(11) DEFAULT NULL,
  `H10` int(11) DEFAULT NULL,
  `H11` int(11) DEFAULT NULL,
  `H12` int(11) DEFAULT NULL,
  `H13` int(11) DEFAULT NULL,
  `H14` int(11) DEFAULT NULL,
  `H15` int(11) DEFAULT NULL,
  `H16` int(11) DEFAULT NULL,
  `H17` int(11) DEFAULT NULL,
  `H18` int(11) DEFAULT NULL,
  `H19` int(11) DEFAULT NULL,
  `H20` int(11) DEFAULT NULL,
  `H21` int(11) DEFAULT NULL,
  `H22` int(11) DEFAULT NULL,
  `H23` int(11) DEFAULT NULL,
  `H24` int(11) DEFAULT NULL,
  `H25` int(11) DEFAULT NULL,
  `H26` int(11) DEFAULT NULL,
  `H27` int(11) DEFAULT NULL,
  `H28` int(11) DEFAULT NULL,
  `H29` int(11) DEFAULT NULL,
  `H30` int(11) DEFAULT NULL,
  `H31` int(11) DEFAULT NULL,
  `H32` int(11) DEFAULT NULL,
  `H33` int(11) DEFAULT NULL,
  `H34` int(11) DEFAULT NULL,
  `H35` int(11) DEFAULT NULL,
  `InvObj1` int(11) DEFAULT NULL,
  `InvCant1` int(11) DEFAULT NULL,
  `InvEqp1` tinyint(1) DEFAULT NULL,
  `InvObj2` int(11) DEFAULT NULL,
  `InvCant2` int(11) DEFAULT NULL,
  `InvEqp2` tinyint(1) DEFAULT NULL,
  `InvObj3` int(11) DEFAULT NULL,
  `InvCant3` int(11) DEFAULT NULL,
  `InvEqp3` tinyint(1) DEFAULT NULL,
  `InvObj4` int(11) DEFAULT NULL,
  `InvCant4` int(11) DEFAULT NULL,
  `InvEqp4` tinyint(1) DEFAULT NULL,
  `InvObj5` int(11) DEFAULT NULL,
  `InvCant5` int(11) DEFAULT NULL,
  `InvEqp5` tinyint(1) DEFAULT NULL,
  `InvObj6` int(11) DEFAULT NULL,
  `InvCant6` int(11) DEFAULT NULL,
  `InvEqp6` tinyint(1) DEFAULT NULL,
  `InvObj7` int(11) DEFAULT NULL,
  `InvCant7` int(11) DEFAULT NULL,
  `InvEqp7` tinyint(1) DEFAULT NULL,
  `InvObj8` int(11) DEFAULT NULL,
  `InvCant8` int(11) DEFAULT NULL,
  `InvEqp8` tinyint(1) DEFAULT NULL,
  `InvObj9` int(11) DEFAULT NULL,
  `InvCant9` int(11) DEFAULT NULL,
  `InvEqp9` tinyint(1) DEFAULT NULL,
  `InvObj10` int(11) DEFAULT NULL,
  `InvCant10` int(11) DEFAULT NULL,
  `InvEqp10` tinyint(1) DEFAULT NULL,
  `InvObj11` int(11) DEFAULT NULL,
  `InvCant11` int(11) DEFAULT NULL,
  `InvEqp11` tinyint(1) DEFAULT NULL,
  `InvObj12` int(11) DEFAULT NULL,
  `InvCant12` int(11) DEFAULT NULL,
  `InvEqp12` tinyint(1) DEFAULT NULL,
  `InvObj13` int(11) DEFAULT NULL,
  `InvCant13` int(11) DEFAULT NULL,
  `InvEqp13` tinyint(1) DEFAULT NULL,
  `InvObj14` int(11) DEFAULT NULL,
  `InvCant14` int(11) DEFAULT NULL,
  `InvEqp14` tinyint(1) DEFAULT NULL,
  `InvObj15` int(11) DEFAULT NULL,
  `InvCant15` int(11) DEFAULT NULL,
  `InvEqp15` tinyint(1) DEFAULT NULL,
  `InvObj16` int(11) DEFAULT NULL,
  `InvCant16` int(11) DEFAULT NULL,
  `InvEqp16` tinyint(1) DEFAULT NULL,
  `InvObj17` int(11) DEFAULT NULL,
  `InvCant17` int(11) DEFAULT NULL,
  `InvEqp17` tinyint(1) DEFAULT NULL,
  `InvObj18` int(11) DEFAULT NULL,
  `InvCant18` int(11) DEFAULT NULL,
  `InvEqp18` tinyint(1) DEFAULT NULL,
  `InvObj19` int(11) DEFAULT NULL,
  `InvCant19` int(11) DEFAULT NULL,
  `InvEqp19` tinyint(1) DEFAULT NULL,
  `InvObj20` int(11) DEFAULT NULL,
  `InvCant20` int(11) DEFAULT NULL,
  `InvEqp20` tinyint(1) DEFAULT NULL,
  `InvCantidadItems` int(11) DEFAULT NULL,
  `GuildIndex` int(11) DEFAULT '0',
  `Descripcion` varchar(100) DEFAULT NULL,
  `Creado` datetime DEFAULT NULL,
  `BannedBy` varchar(20) DEFAULT NULL,
  `Voto` int(11) DEFAULT '0',
  `AspiranteA` int(11) DEFAULT '0',
  `MotivoRechazo` varchar(100) DEFAULT '',
  `Pedidos` varchar(400) DEFAULT '0',
  `Miembro` varchar(400) DEFAULT NULL,
  `Extra` varchar(100) DEFAULT NULL,
  `Penas` int(11) DEFAULT '0',
  `BanTime` date DEFAULT '2000-01-01',
  `Q1` varchar(20) DEFAULT '0',
  `Q2` varchar(20) DEFAULT '0',
  `Q3` varchar(20) DEFAULT '0',
  `Q4` varchar(20) DEFAULT '0',
  `Q5` varchar(20) DEFAULT '0',
  `Q6` varchar(20) DEFAULT '0',
  `Q7` varchar(20) DEFAULT '0',
  `Q8` varchar(20) DEFAULT '0',
  `Q9` varchar(20) DEFAULT '0',
  `Q10` varchar(20) DEFAULT '0',
  `Q11` varchar(20) DEFAULT '0',
  `Q12` varchar(20) DEFAULT '0',
  `Q13` varchar(20) DEFAULT '0',
  `Q14` varchar(20) DEFAULT '0',
  `Q15` varchar(20) DEFAULT '0',
  `QuestsDone` varchar(20) DEFAULT '0'
) ENGINE=MyISAM AUTO_INCREMENT=2364 DEFAULT CHARSET=utf8;

--
-- Volcado de datos para la tabla `pjs`
--

INSERT INTO `pjs` (`Id`, `IdAccount`, `Nombre`, `Genero`, `Raza`, `Hogar`, `Clase`, `Heading`, `Head`, `Body`, `Arma`, `Escudo`, `Casco`, `Uptime`, `LastIP`, `LastConnect`, `Map`, `X`, `Y`, `Logged`, `Muerto`, `Escondido`, `Hambre`, `Sed`, `Desnudo`, `Ban`, `Navegando`, `Envenenado`, `Paralizado`, `PerteneceReal`, `PerteneceCaos`, `Pena`, `EjercitoReal`, `EjercitoCaos`, `CiudMatados`, `CrimMatados`, `rArCaos`, `rArReal`, `rExCaos`, `rExReal`, `recCaos`, `recReal`, `Reenlistadas`, `NivelIngreso`, `FechaIngreso`, `MatadosIngreso`, `NextRecompensa`, `At1`, `At2`, `At3`, `At4`, `At5`, `Sk1`, `Sk2`, `Sk3`, `Sk4`, `Sk5`, `Sk6`, `Sk7`, `Sk8`, `Sk9`, `Sk10`, `Sk11`, `Sk12`, `Sk13`, `Sk14`, `Sk15`, `Sk16`, `Sk17`, `Sk18`, `Sk19`, `Sk20`, `Email`, `Gld`, `Banco`, `MaxHP`, `MinHP`, `MaxSTA`, `MinSTA`, `MaxMAN`, `MinMAN`, `MaxHIT`, `MinHIT`, `MaxAGU`, `MinAGU`, `MaxHAM`, `MinHAM`, `SkillPtsLibres`, `EXP`, `ELV`, `ELU`, `UserMuertes`, `NpcsMuertes`, `WeaponEqpSlot`, `ArmourEqpSlot`, `CascoEqpSlot`, `EscudoEqpSlot`, `BarcoSlot`, `MunicionSlot`, `AnilloSlot`, `Rep_Asesino`, `Rep_Bandido`, `Rep_Burguesia`, `Rep_Ladrones`, `Rep_Nobles`, `Rep_Plebe`, `Rep_Promedio`, `NroMascotas`, `Masc1`, `Masc2`, `Masc3`, `TrainningTime`, `H1`, `H2`, `H3`, `H4`, `H5`, `H6`, `H7`, `H8`, `H9`, `H10`, `H11`, `H12`, `H13`, `H14`, `H15`, `H16`, `H17`, `H18`, `H19`, `H20`, `H21`, `H22`, `H23`, `H24`, `H25`, `H26`, `H27`, `H28`, `H29`, `H30`, `H31`, `H32`, `H33`, `H34`, `H35`, `InvObj1`, `InvCant1`, `InvEqp1`, `InvObj2`, `InvCant2`, `InvEqp2`, `InvObj3`, `InvCant3`, `InvEqp3`, `InvObj4`, `InvCant4`, `InvEqp4`, `InvObj5`, `InvCant5`, `InvEqp5`, `InvObj6`, `InvCant6`, `InvEqp6`, `InvObj7`, `InvCant7`, `InvEqp7`, `InvObj8`, `InvCant8`, `InvEqp8`, `InvObj9`, `InvCant9`, `InvEqp9`, `InvObj10`, `InvCant10`, `InvEqp10`, `InvObj11`, `InvCant11`, `InvEqp11`, `InvObj12`, `InvCant12`, `InvEqp12`, `InvObj13`, `InvCant13`, `InvEqp13`, `InvObj14`, `InvCant14`, `InvEqp14`, `InvObj15`, `InvCant15`, `InvEqp15`, `InvObj16`, `InvCant16`, `InvEqp16`, `InvObj17`, `InvCant17`, `InvEqp17`, `InvObj18`, `InvCant18`, `InvEqp18`, `InvObj19`, `InvCant19`, `InvEqp19`, `InvObj20`, `InvCant20`, `InvEqp20`, `InvCantidadItems`, `GuildIndex`, `Descripcion`, `Creado`, `BannedBy`, `Voto`, `AspiranteA`, `MotivoRechazo`, `Pedidos`, `Miembro`, `Extra`, `Penas`, `BanTime`, `Q1`, `Q2`, `Q3`, `Q4`, `Q5`, `Q6`, `Q7`, `Q8`, `Q9`, `Q10`, `Q11`, `Q12`, `Q13`, `Q14`, `Q15`, `QuestsDone`) VALUES
(2355, 1325, 'deidhara', 1, 2, 1, 7, 3, 110, 210, 15, 2, 2, 26415, '168.196.245.137', '2021-04-26 18:32:55', 1, 231, 1267, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 0, 0, 0, 49, 0, 0, 0, 0, 0, 7, 0, 0, 0, 0, 0, 0, 0, 0, 4, 20, '', 7629, 0, 259, 259, 525, 525, 1290, 50, 64, 63, 100, 100, 100, 100, 150, 122310, 32, 3619679, 0, 72, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 37000, 30, 6172, 0, 0, 0, 0, 7742, 2, 46, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 556, 1, 1, 26, 9, 0, 43, 4, 0, 127, 14, 0, 42, 17, 0, 414, 3, 0, 58, 184, 0, 2, 6, 0, 491, 100, 0, 461, 100, 0, 462, 50, 0, 166, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 12, 0, '', '2021-04-21 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '15-0', '7', '14', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '-2'),
(2356, 1326, 'natalia', 2, 1, 2, 6, 3, 70, 1, 12, 2, 2, 73, '168.196.245.137', '2021-04-24 17:03:20', 1, 179, 1238, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 19, 19, 60, 60, 50, 50, 2, 1, 100, 100, 100, 100, 10, 0, 1, 300, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 72, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 463, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-04-21 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2357, 1325, 'feoo', 1, 1, 1, 4, 3, 22, 1, 12, 2, 2, 60, '168.196.245.137', '2021-04-23 13:16:49', 1, 304, 857, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 19, 19, 40, 40, 50, 50, 2, 1, 100, 100, 100, 100, 10, 0, 1, 300, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 60, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 463, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2358, 1325, 'benjamin', 1, 2, 2, 2, 3, 108, 2, 12, 2, 2, 95, '168.196.245.137', '2021-04-26 13:18:30', 1, 231, 1266, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 100, 0, 21, 21, 60, 60, 50, 50, 2, 1, 100, 100, 100, 100, 10, 150, 1, 300, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 94, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 464, 1, 1, 127, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 7, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2359, 1327, 'pochola', 1, 1, 1, 2, 3, 24, 1, 12, 2, 2, 25, '168.196.245.137', '2021-04-23 11:04:03', 1, 304, 854, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 18, 18, 40, 40, 50, 50, 2, 1, 100, 100, 100, 100, 10, 0, 1, 300, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 25, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 463, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2360, 1325, 'RENATA', 2, 2, 2, 2, 3, 177, 2, 12, 2, 2, 75, '168.196.245.137', '2021-04-23 13:05:25', 1, 229, 1266, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 21, 21, 40, 40, 50, 50, 2, 1, 100, 100, 100, 100, 10, 0, 1, 300, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 75, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 464, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2361, 1326, 'BENJAMINES', 1, 4, 2, 2, 3, 403, 300, 12, 2, 2, 234, '168.196.245.137', '2021-04-23 13:21:32', 1, 231, 1266, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 16, 21, 22, 19, 18, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 100, 0, 19, 19, 40, 40, 50, 50, 2, 1, 100, 100, 100, 100, 10, 150, 1, 300, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 233, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 466, 1, 1, 127, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 7, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1-1'),
(2362, 1328, 'sensui', 1, 1, 2, 2, 2, 37, 3, 15, 2, 2, 1349, '213.94.53.229', '2021-04-23 16:28:05', 1, 226, 1269, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 22, '', 132475, 0, 288, 288, 555, 555, 1238, 50, 68, 67, 100, 100, 100, 100, 158, 1552744, 34, 6596865, 0, 5, 4, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3500, 30, 588, 0, 0, 0, 0, 1081, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 389, 1, 0, 2, 1, 0, 35, 1, 1, 556, 1, 1, 1100, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 5, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '3-7-14-15'),
(2363, 1328, 'asasa', 1, 1, 2, 2, 1, 19, 21, 2, 2, 2, 496, '213.94.53.229', '2021-04-23 16:35:38', 1, 223, 1266, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 257, 0, 505, 0, 1166, 50, 64, 63, 100, 100, 100, 100, 170, 2925510, 32, 3619679, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2000, 32, 339, 0, 0, 0, 0, 382, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '15-2', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `quest`
--

CREATE TABLE IF NOT EXISTS `quest` (
  `user_id` int(11) NOT NULL,
  `number` int(11) NOT NULL,
  `quest_id` int(11) NOT NULL,
  `npcs` varchar(64) NOT NULL,
  `npcstarget` varchar(64) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Volcado de datos para la tabla `quest`
--

INSERT INTO `quest` (`user_id`, `number`, `quest_id`, `npcs`, `npcstarget`) VALUES
(2355, 1, 0, '', ''),
(2355, 2, 0, '', ''),
(2355, 3, 0, '', ''),
(2355, 4, 0, '', ''),
(2355, 5, 0, '', ''),
(2358, 1, 0, '', ''),
(2358, 2, 0, '', ''),
(2358, 3, 0, '', ''),
(2358, 4, 0, '', ''),
(2358, 5, 0, '', '');

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `quest_done`
--

CREATE TABLE IF NOT EXISTS `quest_done` (
  `user_id` mediumint(8) unsigned NOT NULL,
  `quest_id` tinyint(3) unsigned NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Volcado de datos para la tabla `quest_done`
--

INSERT INTO `quest_done` (`user_id`, `quest_id`) VALUES
(2355, 0),
(2355, 2),
(2355, 3);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `vault`
--

CREATE TABLE IF NOT EXISTS `vault` (
  `cuenta_id` bigint(20) NOT NULL,
  `slot` int(11) NOT NULL,
  `item` int(11) NOT NULL DEFAULT '0',
  `quantity` int(11) NOT NULL DEFAULT '0',
  `updated_at` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `clanes`
--
ALTER TABLE `clanes`
  ADD PRIMARY KEY (`Id`),
  ADD UNIQUE KEY `Id` (`Id`);

--
-- Indices de la tabla `clanes_propuestas`
--
ALTER TABLE `clanes_propuestas`
  ADD PRIMARY KEY (`Id`),
  ADD UNIQUE KEY `Id` (`Id`);

--
-- Indices de la tabla `clanes_relaciones`
--
ALTER TABLE `clanes_relaciones`
  ADD PRIMARY KEY (`IdClan`,`IdClanTo`);

--
-- Indices de la tabla `clanes_solicitudes`
--
ALTER TABLE `clanes_solicitudes`
  ADD PRIMARY KEY (`Id`),
  ADD UNIQUE KEY `Id` (`Id`);

--
-- Indices de la tabla `clanes_votos`
--
ALTER TABLE `clanes_votos`
  ADD PRIMARY KEY (`IdClan`,`Nombre`);

--
-- Indices de la tabla `conexiones`
--
ALTER TABLE `conexiones`
  ADD PRIMARY KEY (`Id`);

--
-- Indices de la tabla `cuentas`
--
ALTER TABLE `cuentas`
  ADD PRIMARY KEY (`Id`),
  ADD UNIQUE KEY `Nombre` (`Nombre`);

--
-- Indices de la tabla `fortalezas`
--
ALTER TABLE `fortalezas`
  ADD PRIMARY KEY (`Id`),
  ADD UNIQUE KEY `Id` (`Id`);

--
-- Indices de la tabla `penas`
--
ALTER TABLE `penas`
  ADD PRIMARY KEY (`Id`),
  ADD UNIQUE KEY `Id` (`Id`);

--
-- Indices de la tabla `pjs`
--
ALTER TABLE `pjs`
  ADD PRIMARY KEY (`Id`),
  ADD UNIQUE KEY `Nombre` (`Nombre`);

--
-- Indices de la tabla `quest`
--
ALTER TABLE `quest`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `quest_done`
--
ALTER TABLE `quest_done`
  ADD PRIMARY KEY (`user_id`,`quest_id`);

--
-- Indices de la tabla `vault`
--
ALTER TABLE `vault`
  ADD PRIMARY KEY (`cuenta_id`,`slot`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `clanes`
--
ALTER TABLE `clanes`
  MODIFY `Id` int(11) NOT NULL AUTO_INCREMENT,AUTO_INCREMENT=13;
--
-- AUTO_INCREMENT de la tabla `clanes_propuestas`
--
ALTER TABLE `clanes_propuestas`
  MODIFY `Id` bigint(20) NOT NULL AUTO_INCREMENT;
--
-- AUTO_INCREMENT de la tabla `clanes_solicitudes`
--
ALTER TABLE `clanes_solicitudes`
  MODIFY `Id` bigint(20) NOT NULL AUTO_INCREMENT,AUTO_INCREMENT=159;
--
-- AUTO_INCREMENT de la tabla `conexiones`
--
ALTER TABLE `conexiones`
  MODIFY `Id` bigint(20) NOT NULL AUTO_INCREMENT;
--
-- AUTO_INCREMENT de la tabla `cuentas`
--
ALTER TABLE `cuentas`
  MODIFY `Id` bigint(20) NOT NULL AUTO_INCREMENT,AUTO_INCREMENT=1329;
--
-- AUTO_INCREMENT de la tabla `fortalezas`
--
ALTER TABLE `fortalezas`
  MODIFY `Id` int(11) NOT NULL AUTO_INCREMENT,AUTO_INCREMENT=3;
--
-- AUTO_INCREMENT de la tabla `penas`
--
ALTER TABLE `penas`
  MODIFY `Id` tinyint(20) NOT NULL AUTO_INCREMENT,AUTO_INCREMENT=128;
--
-- AUTO_INCREMENT de la tabla `pjs`
--
ALTER TABLE `pjs`
  MODIFY `Id` bigint(20) unsigned NOT NULL AUTO_INCREMENT,AUTO_INCREMENT=2364;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
