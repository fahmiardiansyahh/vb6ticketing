-- phpMyAdmin SQL Dump
-- version 4.3.11
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: May 09, 2017 at 02:26 PM
-- Server version: 5.6.24
-- PHP Version: 5.6.8

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `db_tiket`
--

-- --------------------------------------------------------

--
-- Table structure for table `tb_admin`
--

CREATE TABLE IF NOT EXISTS `tb_admin` (
  `id_petugas` varchar(10) NOT NULL,
  `nama_petugas` varchar(30) NOT NULL,
  `password` varchar(8) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_admin`
--

INSERT INTO `tb_admin` (`id_petugas`, `nama_petugas`, `password`) VALUES
('PT001', 'VIKRIAWAN', '12345678');

-- --------------------------------------------------------

--
-- Table structure for table `tb_detailtransaksi`
--

CREATE TABLE IF NOT EXISTS `tb_detailtransaksi` (
  `notransaksi` varchar(12) NOT NULL,
  `jumlah_pesanan` int(11) NOT NULL,
  `sub_total` double NOT NULL,
  `id_kereta` char(9) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_detailtransaksi`
--

INSERT INTO `tb_detailtransaksi` (`notransaksi`, `jumlah_pesanan`, `sub_total`, `id_kereta`) VALUES
('TR-17050701', 1, 100000, 'FU10G400P'),
('TR-17050702', 1, 70000, 'MJ10G400P'),
('TR-17050703', 2, 200000, 'FU10G400P'),
('TR-17050704', 2, 200000, 'FU10G400P'),
('TR-17050901', 2, 200000, 'FU10G400P');

-- --------------------------------------------------------

--
-- Table structure for table `tb_kereta`
--

CREATE TABLE IF NOT EXISTS `tb_kereta` (
  `id_kereta` varchar(9) NOT NULL,
  `nama_kereta` varchar(20) NOT NULL,
  `pemberangkatan` varchar(15) NOT NULL,
  `tujuan` char(15) NOT NULL,
  `kelas` char(10) NOT NULL,
  `kursi_tersedia` int(11) NOT NULL,
  `stasiun_keberangkatan` varchar(20) NOT NULL,
  `tanggal_keberangkatan` date NOT NULL,
  `jam_keberangkatan` varchar(5) NOT NULL,
  `stasiun_tujuan` varchar(20) NOT NULL,
  `tanggal_tiba` date NOT NULL,
  `jam_tiba` varchar(5) NOT NULL,
  `harga_tiket` int(10) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_kereta`
--

INSERT INTO `tb_kereta` (`id_kereta`, `nama_kereta`, `pemberangkatan`, `tujuan`, `kelas`, `kursi_tersedia`, `stasiun_keberangkatan`, `tanggal_keberangkatan`, `jam_keberangkatan`, `stasiun_tujuan`, `tanggal_tiba`, `jam_tiba`, `harga_tiket`) VALUES
('FU10G400P', 'KA.FAJARUTAMA', 'YOGYAKARTA', 'JAKARTA', 'BISNIS', 379, 'ST.TUGU', '2017-08-16', '07:00', 'PS.SENEN', '2017-08-19', '17:00', 100000),
('GY07G175P', 'KA.GAJAYANA', 'MALANG', 'JAKARTA', 'EKSEKUTIF', 170, 'ST.MALANG', '2017-08-21', '23:00', 'ST.GAMBIR', '2017-08-23', '01:00', 200000),
('JT11G550P', 'KA.JAKATINGKIR', 'SURAKARTA', 'JAKARTA', 'EKONOMI AC', 596, 'ST.PURWOSARI', '2017-08-16', '02:00', 'PS.SENEN', '2017-08-18', '11:00', 60000),
('KM13G133P', 'KA.KERETAMALAM', 'SEMARANG', 'JAKARTA', 'EKSEKUTIF', 129, 'ST.LAWANGSEWU', '2017-08-28', '24:00', 'ST.GAMBIR', '2017-08-31', '00:00', 300000),
('KT12G720P', 'KA.KRAKATAU', 'BANTEN', 'BLITAR', 'EKONOMI AC', 718, 'MERAK', '2017-08-22', '05:00', 'BLITAR', '2017-08-24', '07:00', 65000),
('MJ10G400P', 'KA.MAJAPAHIT', 'JAKARTA', 'MALANG', 'EKONOMI AC', 394, 'PS.SENEN', '2017-08-11', '03:00', 'ST.MALANG', '2017-08-13', '09:00', 70000),
('MS10G450P', 'KA.MUTIARASELATAN', 'BANDUNG', 'MALANG', 'BISNIS', 450, 'ST.BANDUNG', '2017-08-02', '02:00', 'ST.MALANG', '2017-08-03', '11:30', 50000),
('SB06G180P', 'KA.SEMBRANI', 'SURABAYA', 'JAKARTA', 'EKSEKUTIF', 178, 'ST.PASARTURI', '2017-09-01', '14:00', 'ST.GAMBIR', '2017-09-04', '10:30', 120000),
('SM09G315P', 'KA.SIDOMUKTI', 'SOLO', 'YOGYAKARTA', 'BISNIS', 313, 'ST.BALAPAN', '2017-09-04', '08:00', 'ST.TUGU', '2017-09-04', '20:00', 80000);

-- --------------------------------------------------------

--
-- Table structure for table `tb_pemesan`
--

CREATE TABLE IF NOT EXISTS `tb_pemesan` (
  `id_pemesan` char(10) NOT NULL,
  `nama_pemesan` varchar(30) NOT NULL,
  `password` varchar(8) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_pemesan`
--

INSERT INTO `tb_pemesan` (`id_pemesan`, `nama_pemesan`, `password`) VALUES
('PN001', 'FAHMI ARDIANSYAH', '12345678');

-- --------------------------------------------------------

--
-- Table structure for table `tb_temp`
--

CREATE TABLE IF NOT EXISTS `tb_temp` (
  `notransaksi` varchar(11) NOT NULL,
  `tgl_pesan` date NOT NULL,
  `nama_pemesan` varchar(30) NOT NULL,
  `kode_kereta` char(9) NOT NULL,
  `nama_kereta` varchar(20) NOT NULL,
  `kelas` char(10) NOT NULL,
  `dari` varchar(20) NOT NULL,
  `tujuan` varchar(20) NOT NULL,
  `jumlah_pesan` int(11) NOT NULL,
  `harga` int(11) NOT NULL,
  `subtotal` int(11) NOT NULL,
  `uang_bayar` int(11) NOT NULL,
  `kembali` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `tb_transaksi`
--

CREATE TABLE IF NOT EXISTS `tb_transaksi` (
  `notransaksi` varchar(12) NOT NULL,
  `tgl_jual` date NOT NULL,
  `id_pemesan` char(5) NOT NULL,
  `nama_pemesan` varchar(30) NOT NULL,
  `id_kereta` char(9) NOT NULL,
  `nama_kereta` varchar(20) NOT NULL,
  `pemberangkatan` varchar(15) NOT NULL,
  `tujuan` varchar(15) NOT NULL,
  `kelas` varchar(10) NOT NULL,
  `harga` int(11) NOT NULL,
  `jumlah_beli` int(11) NOT NULL,
  `sub_total` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_transaksi`
--

INSERT INTO `tb_transaksi` (`notransaksi`, `tgl_jual`, `id_pemesan`, `nama_pemesan`, `id_kereta`, `nama_kereta`, `pemberangkatan`, `tujuan`, `kelas`, `harga`, `jumlah_beli`, `sub_total`) VALUES
('TR-17050701', '2017-05-07', 'PN001', 'FAHMI ARDIANSYAH', 'FU10G400P', 'KA.FAJARUTAMA', 'YOGYAKARTA', 'JAKARTA', 'BISNIS', 100000, 1, 100000),
('TR-17050702', '2017-05-07', 'PN001', 'FAHMI ARDIANSYAH', 'MJ10G400P', 'KA.MAJAPAHIT', 'JAKARTA', 'MALANG', 'EKONOMI AC', 70000, 1, 70000),
('TR-17050703', '2017-05-07', 'PN001', 'FAHMI ARDIANSYAH', 'FU10G400P', 'KA.FAJARUTAMA', 'YOGYAKARTA', 'JAKARTA', 'BISNIS', 100000, 2, 200000),
('TR-17050704', '2017-05-07', 'PN001', 'FAHMI ARDIANSYAH', 'FU10G400P', 'KA.FAJARUTAMA', 'YOGYAKARTA', 'JAKARTA', 'BISNIS', 100000, 2, 200000),
('TR-17050901', '2017-05-09', 'PN001', 'FAHMI ARDIANSYAH', 'FU10G400P', 'KA.FAJARUTAMA', 'YOGYAKARTA', 'JAKARTA', 'BISNIS', 100000, 2, 200000);

--
-- Indexes for dumped tables
--

--
-- Indexes for table `tb_kereta`
--
ALTER TABLE `tb_kereta`
  ADD PRIMARY KEY (`id_kereta`);

--
-- Indexes for table `tb_pemesan`
--
ALTER TABLE `tb_pemesan`
  ADD PRIMARY KEY (`id_pemesan`);

--
-- Indexes for table `tb_transaksi`
--
ALTER TABLE `tb_transaksi`
  ADD PRIMARY KEY (`notransaksi`);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
