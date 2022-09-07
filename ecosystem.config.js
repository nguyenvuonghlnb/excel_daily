module.exports = {
  apps : [
      {
        name: "project_excel",
        script: "./main.py",
        watch: false,
        instances: 1,
        cron_restart: "40 19 * * *",
        stop_exit_codes: [0],
        env: {
            app_env: "production"
        }
      }
  ]
}