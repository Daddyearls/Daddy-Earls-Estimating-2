# Daddy Earl's Estimating

A lightweight estimating intake and owner dashboard for Daddy Earl's Auto Repair.

## Features
- Public customer estimate request form
- Private owner sign-in and request dashboard
- Automotive, light-duty, medium-duty, heavy-duty, commercial, marine, and off-road vehicle categories
- Estimate builder with a $135/hour labor rate and 40% parts markup
- OEM, aftermarket, and customer-supplied parts choices
- Owner-controlled estimate approval and request status tracking
- PostgreSQL persistence

## Required environment variables
- `DATABASE_URL`: Render PostgreSQL internal connection string
- `ADMIN_PASSWORD`: strong owner dashboard password
- `SESSION_SECRET`: long random secret used to sign sessions
- `NODE_ENV=production`

## Local development
```sh
npm install
DATABASE_URL=postgres://... ADMIN_PASSWORD=... SESSION_SECRET=... npm start
```

The public estimate request page is at `/`; the private owner dashboard is at `/dashboard`.
