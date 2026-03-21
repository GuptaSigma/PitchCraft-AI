import os

import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2.pool import ThreadedConnectionPool


connection_pool = None


class PooledConnection:
    """Small wrapper to keep MySQL-like cursor(dictionary=True) compatibility."""

    def __init__(self, pool, conn):
        self._pool = pool
        self._conn = conn

    def cursor(self, *args, **kwargs):
        dictionary = kwargs.pop('dictionary', False)
        if dictionary:
            kwargs['cursor_factory'] = RealDictCursor
        return self._conn.cursor(*args, **kwargs)

    def __getattr__(self, name):
        return getattr(self._conn, name)

    def close(self):
        self._pool.putconn(self._conn)


def get_database_url():
    database_url = os.getenv('DATABASE_URL', '').strip()
    if database_url:
        return database_url

    host = os.getenv('DB_HOST', 'localhost')
    user = os.getenv('DB_USER', 'postgres')
    password = os.getenv('DB_PASSWORD', '')
    db_name = os.getenv('DB_NAME', 'gamma_ai')
    port = os.getenv('DB_PORT', '5432')
    return f"postgresql://{user}:{password}@{host}:{port}/{db_name}"


def get_connection():
    """Get PostgreSQL connection from pool."""
    global connection_pool

    try:
        if connection_pool is None:
            connection_pool = ThreadedConnectionPool(
                minconn=1,
                maxconn=int(os.getenv('DB_POOL_SIZE', '10')),
                dsn=get_database_url()
            )
        conn = connection_pool.getconn()
        conn.autocommit = False
        return PooledConnection(connection_pool, conn)
    except Exception as err:
        print(f"❌ Error getting PostgreSQL connection: {err}")
        raise


def init_db():
    """Initialize PostgreSQL tables."""

    print("\n" + "=" * 60)
    print("📊 INITIALIZING POSTGRESQL DATABASE")
    print("=" * 60)

    conn = None
    cursor = None

    try:
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id BIGSERIAL PRIMARY KEY,
                name VARCHAR(255) NOT NULL,
                email VARCHAR(255) UNIQUE NOT NULL,
                password VARCHAR(255) NULL,
                google_id VARCHAR(255) NULL,
                auth_provider VARCHAR(50) NOT NULL DEFAULT 'password',
                email_verified BOOLEAN NOT NULL DEFAULT FALSE,
                created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_email ON users (email)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_created_at ON users (created_at)")
        cursor.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_google_id ON users (google_id)")
        print("✅ Table 'users' ready")

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS otp_verifications (
                id BIGSERIAL PRIMARY KEY,
                email VARCHAR(255) NOT NULL,
                name VARCHAR(255) NULL,
                password_hash VARCHAR(255) NULL,
                otp_code VARCHAR(10) NOT NULL,
                purpose VARCHAR(50) NOT NULL,
                expires_at TIMESTAMPTZ NOT NULL,
                consumed_at TIMESTAMPTZ NULL,
                created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        cursor.execute(
            "CREATE INDEX IF NOT EXISTS idx_otp_email_purpose ON otp_verifications (email, purpose)"
        )
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_otp_expiry ON otp_verifications (expires_at)")
        print("✅ Table 'otp_verifications' ready")

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS user_daily_usage (
                id BIGSERIAL PRIMARY KEY,
                user_id BIGINT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                usage_date DATE NOT NULL,
                tokens_used INT NOT NULL DEFAULT 0,
                created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
                CONSTRAINT uniq_user_usage_date UNIQUE (user_id, usage_date)
            )
            """
        )
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_usage_date ON user_daily_usage (usage_date)")
        print("✅ Table 'user_daily_usage' ready")

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS presentations (
                id BIGSERIAL PRIMARY KEY,
                user_id BIGINT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                title VARCHAR(500) NOT NULL,
                prompt TEXT NOT NULL,
                slides_count INT DEFAULT 10,
                total_slides INT DEFAULT 10,
                content JSONB,
                output_type VARCHAR(50) DEFAULT 'presentation',
                style VARCHAR(50) DEFAULT 'business',
                theme VARCHAR(50) DEFAULT 'alien',
                language VARCHAR(10) DEFAULT 'en-uk',
                image_style VARCHAR(50) DEFAULT 'illustration',
                text_amount VARCHAR(20) DEFAULT 'moderate',
                ai_model VARCHAR(100) DEFAULT 'gemini-2.0-flash',
                created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_presentations_user_id ON presentations (user_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_presentations_created_at ON presentations (created_at)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_presentations_output_type ON presentations (output_type)")
        print("✅ Table 'presentations' ready")

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS slides (
                id BIGSERIAL PRIMARY KEY,
                presentation_id BIGINT NOT NULL REFERENCES presentations(id) ON DELETE CASCADE,
                slide_number INT NOT NULL,
                title VARCHAR(500),
                content TEXT,
                layout JSONB,
                image_url VARCHAR(1000),
                background VARCHAR(500),
                animation VARCHAR(50),
                notes TEXT,
                metadata JSONB,
                created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_slides_presentation_id ON slides (presentation_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_slides_slide_number ON slides (slide_number)")
        print("✅ Table 'slides' ready")

        conn.commit()
        print("=" * 60)
        print("✅ PostgreSQL tables created successfully!")
        print("=" * 60 + "\n")

    except Exception as err:
        print(f"\n❌ Database error: {err}")
        if conn:
            conn.rollback()
        raise

    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def execute_query(query, params=None, fetch=False):
    """Execute database query with PostgreSQL compatibility."""

    conn = None
    cursor = None

    try:
        conn = get_connection()
        cursor = conn.cursor(dictionary=True)

        normalized = (query or '').strip()
        is_insert = normalized.lower().startswith('insert')

        if is_insert and 'returning' not in normalized.lower() and not fetch:
            returning_query = normalized.rstrip(';') + ' RETURNING id'
            cursor.execute(returning_query, params or ())
            row = cursor.fetchone()
            conn.commit()
            if isinstance(row, dict):
                return row.get('id')
            return row[0] if row else None

        cursor.execute(query, params or ())

        if fetch:
            return cursor.fetchall()

        conn.commit()
        return cursor.rowcount

    except Exception as err:
        print(f"❌ Query error: {err}")
        if conn:
            conn.rollback()
        raise

    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()