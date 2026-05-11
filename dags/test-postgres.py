from airflow.providers.postgres.hooks.postgres import PostgresHook

hook = PostgresHook(postgres_conn_id='gcp-postgres')
conn = hook.get_conn()