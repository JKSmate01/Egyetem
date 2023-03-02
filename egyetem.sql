-- phpMyAdmin SQL Dump
-- version 4.9.0.1
-- https://www.phpmyadmin.net/
--
-- Gép: 127.0.0.1
-- Létrehozás ideje: 2023. Jan 27. 13:26
-- Kiszolgáló verziója: 10.4.6-MariaDB
-- PHP verzió: 7.3.8

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Adatbázis: `egyetem`
--

-- --------------------------------------------------------

--
-- Tábla szerkezet ehhez a táblához `eloadasok`
--

CREATE TABLE `eloadasok` (
  `id` int(11) NOT NULL,
  `cim` varchar(50) COLLATE utf8_hungarian_ci NOT NULL,
  `idopont` varchar(50) COLLATE utf8_hungarian_ci NOT NULL,
  `tanar_id` int(11) NOT NULL,
  `tanfelev_id` int(11) NOT NULL,
  `tanszek_id` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_hungarian_ci;

-- --------------------------------------------------------

--
-- Tábla szerkezet ehhez a táblához `tanarok`
--

CREATE TABLE `tanarok` (
  `id` int(11) NOT NULL,
  `nev` varchar(40) COLLATE utf8_hungarian_ci NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_hungarian_ci;

-- --------------------------------------------------------

--
-- Tábla szerkezet ehhez a táblához `tanfelev`
--

CREATE TABLE `tanfelev` (
  `id` int(11) NOT NULL,
  `tanfelev` varchar(12) COLLATE utf8_hungarian_ci DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_hungarian_ci;

-- --------------------------------------------------------

--
-- Tábla szerkezet ehhez a táblához `tanszek`
--

CREATE TABLE `tanszek` (
  `id` int(11) NOT NULL,
  `nev` varchar(255) COLLATE utf8_hungarian_ci NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_hungarian_ci;

--
-- Indexek a kiírt táblákhoz
--

--
-- A tábla indexei `eloadasok`
--
ALTER TABLE `eloadasok`
  ADD PRIMARY KEY (`id`),
  ADD KEY `eloadas_tan` (`tanar_id`),
  ADD KEY `tanfelev_index` (`tanfelev_id`);

--
-- A tábla indexei `tanarok`
--
ALTER TABLE `tanarok`
  ADD PRIMARY KEY (`id`),
  ADD KEY `tanarindex` (`id`);

--
-- A tábla indexei `tanfelev`
--
ALTER TABLE `tanfelev`
  ADD PRIMARY KEY (`id`);

--
-- A tábla indexei `tanszek`
--
ALTER TABLE `tanszek`
  ADD PRIMARY KEY (`id`),
  ADD KEY `id_index` (`id`);

--
-- Megkötések a kiírt táblákhoz
--

--
-- Megkötések a táblához `eloadasok`
--
ALTER TABLE `eloadasok`
  ADD CONSTRAINT `eloadasok_ibfk_1` FOREIGN KEY (`tanar_id`) REFERENCES `tanarok` (`id`) ON DELETE CASCADE ON UPDATE CASCADE,
  ADD CONSTRAINT `eloadasok_ibfk_2` FOREIGN KEY (`tanfelev_id`) REFERENCES `tanfelev` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
