{
  "version": 2,
  "builds": [
    { "src": "api/processar.js", "use": "@vercel/node" }
  ],
  "routes": [
    { "src": "/api/(.*)", "dest": "/api/processar.js" }
  ]
}
