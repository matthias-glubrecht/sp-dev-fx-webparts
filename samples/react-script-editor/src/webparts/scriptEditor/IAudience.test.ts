import { toAudience, type ISharePointPickerEntity } from './IAudience';
const context = {
    pageContext: {
        web: {
            absoluteUrl: 'https://contoso.sharepoint.com/sites/demo'
        }
    }
};

describe('toAudience', () => {
    it('preserves the stored user audience shape', () => {
        const entity: ISharePointPickerEntity = {
            Key: 'i:0#.f|membership|user@contoso.com',
            DisplayText: 'Sample User',
            Description: 'user@contoso.com',
            EntityType: 'User',
            EntityData: {
                Email: 'user@contoso.com',
                Title: 'Engineer'
            }
        };

        expect(toAudience(entity, context)).toEqual({
            id: 'i:0#.f|membership|user@contoso.com',
            fullName: 'Sample User',
            login: 'user@contoso.com',
            email: 'user@contoso.com',
            jobTitle: 'Engineer',
            initials: 'SU',
            imageUrl: 'https://contoso.sharepoint.com/sites/demo/_layouts/15/userphoto.aspx?size=S&accountname=user%40contoso.com'
        });
    });

    it('preserves SharePoint group IDs and Entra group claims', () => {
        const sharePointGroup: ISharePointPickerEntity = {
            Key: '7',
            DisplayText: 'Site Members',
            Description: 'Members of the site',
            EntityType: 'SPGroup',
            EntityData: { SPGroupID: '7', AccountName: 'Site Members' }
        };
        const entraGroup: ISharePointPickerEntity = {
            Key: 'c:0o.c|federateddirectoryclaimprovider|group-id',
            DisplayText: 'Entra Group',
            Description: 'Entra security group',
            EntityType: 'SecGroup',
            ProviderName: 'FederatedDirectoryClaimProvider'
        };

        expect(toAudience(sharePointGroup, context)).toEqual({
            id: '7',
            fullName: 'Site Members',
            login: 'Site Members',
            description: 'Members of the site'
        });
        expect(toAudience(entraGroup, context)).toEqual({
            id: 'c:0o.c|federateddirectoryclaimprovider|group-id',
            fullName: 'Entra Group',
            login: 'FederatedDirectoryClaimProvider',
            description: 'Entra security group'
        });
    });

    it('ignores malformed search results', () => {
        expect(toAudience({
            DisplayText: 'Missing key',
            Description: 'No key',
            EntityType: 'User',
            EntityData: {}
        }, context)).toBeUndefined();
    });
});
