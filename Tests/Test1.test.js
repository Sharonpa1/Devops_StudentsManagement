const request = require('supertest')
const server = require('../server')

describe("Test suite 1:", ()=>{
    test("test 1: ", async ()=>{
        const res = await request(server).get('/')
        expect(res.statusCode).toEqual(200)
    })
    test("test 2: ", async ()=>{
        const res = await request(server).get('/1111')
        expect(res.statusCode).toEqual(404)
    })
})
