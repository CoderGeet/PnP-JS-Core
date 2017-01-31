import { IConfigurationProvider } from "../configuration";
import { TypedHash } from "../../collections/collections";
import { default as CachingConfigurationProvider } from "./cachingConfigurationProvider";
<<<<<<< HEAD
import { Web } from "../../sharepoint/rest/webs";
import * as Util from "../../utils/util";
=======
import { Web } from "../../sharepoint/webs";
>>>>>>> 159efbbcd163ad9395711d664f9e2121adf27b86

/**
 * A configuration provider which loads configuration values from a SharePoint list
 *
 */
export default class SPListConfigurationProvider implements IConfigurationProvider {
    /**
     * Creates a new SharePoint list based configuration provider
     * @constructor
     * @param {string} webUrl Url of the SharePoint site, where the configuration list is located
     * @param {string} listTitle Title of the SharePoint list, which contains the configuration settings (optional, default = "config")
     */
    constructor(private sourceWeb: Web, private sourceListTitle = "config") {
    }

    /**
     * Gets the url of the SharePoint site, where the configuration list is located
     *
     * @return {string} Url address of the site
     */
    public get web(): Web {
        return this.sourceWeb;
    }

    /**
     * Gets the title of the SharePoint list, which contains the configuration settings
     *
     * @return {string} List title
     */
    public get listTitle(): string {
        return this.sourceListTitle;
    }

    /**
     * Loads the configuration values from the SharePoint list
     *
     * @return {Promise<TypedHash<string>>} Promise of loaded configuration values
     */
    public getConfiguration(): Promise<TypedHash<string>> {

<<<<<<< HEAD
        let web = new Web(Util.combinePaths(this.webUrl, "_api"));
        return web.lists.getByTitle(this.listTitle).items.select("Title", "Value").get().then(function (data) {
            let configuration: TypedHash<string> = {};
            data.forEach(i => {
                configuration[i.Title] = i.Value;
            });
            return configuration;
        });
=======
        return this.web.lists.getByTitle(this.listTitle).items.select("Title", "Value")
            .getAs<{ Title: string, Value: string }[]>().then((data) => {
                return data.reduce((configuration, item) => {

                    return Object.defineProperty(configuration, item.Title, {
                        configurable: false,
                        enumerable: false,
                        value: item.Value,
                        writable: false,
                    });

                }, {});
            });
>>>>>>> 159efbbcd163ad9395711d664f9e2121adf27b86
    }

    /**
     * Wraps the current provider in a cache enabled provider
     *
     * @return {CachingConfigurationProvider} Caching providers which wraps the current provider
     */
    public asCaching(): CachingConfigurationProvider {
<<<<<<< HEAD
        let cacheKey = `splist_${this.webUrl}+${this.listTitle}`;
=======
        let cacheKey = `splist_${this.web.toUrl()}+${this.listTitle}`;
>>>>>>> 159efbbcd163ad9395711d664f9e2121adf27b86
        return new CachingConfigurationProvider(this, cacheKey);
    }
}
