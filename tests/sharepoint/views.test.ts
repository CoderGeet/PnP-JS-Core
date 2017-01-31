import { expect } from "chai";
import { Views } from "../../src/sharepoint/views";
import { toMatchEndRegex } from "../testutils";

describe("Views", () => {
    it("Should be an object", () => {
        let views = new Views("_api/web/lists/getByTitle('Tasks')");
        expect(views).to.be.a("object");
    });
    describe("url", () => {
        it("Should return _api/web/lists/getByTitle('Tasks')/Views", () => {
            let views = new Views("_api/web/lists/getByTitle('Tasks')");
            expect(views.toUrl()).to.match(toMatchEndRegex("_api/web/lists/getByTitle('Tasks')/views"));
        });
    });
    describe("getById", () => {
        it("Should return _api/web/lists/getByTitle('Tasks')/Views('7b7c777e-b749-4f58-a825-53084f2941b0')", () => {
            let views = new Views("_api/web/lists/getByTitle('Tasks')");
            let view = views.getById("7b7c777e-b749-4f58-a825-53084f2941b0");
<<<<<<< HEAD:src/sharepoint/rest/tests/views.test.ts
            expect(view.toUrl()).to.equal("_api/web/lists/getByTitle('Tasks')/views('7b7c777e-b749-4f58-a825-53084f2941b0')");
=======
            expect(view.toUrl())
                .to.match(toMatchEndRegex("_api/web/lists/getByTitle('Tasks')/views('7b7c777e-b749-4f58-a825-53084f2941b0')"));
>>>>>>> 159efbbcd163ad9395711d664f9e2121adf27b86:tests/sharepoint/views.test.ts
        });
    });
});
