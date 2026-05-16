module.exports = {
  apps: [{
    name: 'tele-bot',
    script: 'files/bot.py',
    interpreter: 'python',
    watch: false,
    max_memory_restart: '200M',
    restart_delay: 3000,
    env: {
      PYTHONUNBUFFERED: '1'
    },
    error_file: 'logs/err.log',
    out_file: 'logs/out.log',
    log_date_format: 'YYYY-MM-DD HH:mm:ss'
  }]
};
