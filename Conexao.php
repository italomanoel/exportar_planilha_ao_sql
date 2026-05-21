<?php

class Conexao
{
    private static $conn = null;
    /**
     * Establishes a connection to the SQL Server database.
     * @param string $serverName (e.g., "localhost")
     * @param string $database   (e.g., "db")
     * @param string $username   (e.g., "user")
     * @param string $password   (e.g., "password")
     * @return PDO
     */
    public static function getConnection(string $serverName = "", string $database = "", string $username = "", string $password = "")
    {
        try {
            $dsn = "sqlsrv:server={$serverName};Database={$database}";
            $conn = new PDO($dsn, $username, $password);
            $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
            // Cache the connection for reuse
            self::$conn = $conn;
            return $conn;
        } catch (PDOException $e) {
            die("Erro na conexão: " . $e->getMessage());
        }
    }

    /**
     * Gets the cached database connection.
     *
     * @return PDO|null
     */
    public static function get()
    {
        return self::$conn;
    }

    /**
     * Closes the database connection (optional, but good practice).
     */
    public static function closeConnection()
    {
        if (self::$conn) {
            self::$conn = null; // Disconnects the connection
        }
    }
}
