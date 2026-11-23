export interface IAudience {
    id?: string;
    description?: string;
    fullName: string;
    login: string;
    email?: string;
    jobTitle?: string;
    initials?: string;
    imageUrl?: string;
}

export interface ISharePointPickerEntity {
    Key?: string;
    DisplayText: string;
    Description: string;
    EntityType: string;
    ProviderName?: string;
    EntityData?: {
        Email?: string;
        PrincipalType?: string;
        Title?: string;
        AccountName?: string;
        SPGroupID?: string;
    };
}

export interface IAudiencePickerContext {
    pageContext: {
        web: {
            absoluteUrl: string;
        };
    };
}

export function toAudience(entity: ISharePointPickerEntity, context: IAudiencePickerContext): IAudience | undefined {
    if (!entity.Key) {
        return undefined;
    }

    if (entity.EntityType === 'User') {
        const email = entity.EntityData?.Email || entity.Description;
        return {
            id: entity.Key,
            fullName: entity.DisplayText,
            login: entity.Description,
            email,
            jobTitle: entity.EntityData?.Title,
            initials: entity.DisplayText.split(/\s+/).slice(0, 2).map(part => part.charAt(0)).join(''),
            imageUrl: email
                ? `${context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(email)}`
                : undefined
        };
    }

    if (entity.EntityType === 'SecGroup' || entity.EntityType === 'FormsRole') {
        return {
            id: entity.Key,
            fullName: entity.DisplayText,
            login: entity.ProviderName || '',
            description: entity.Description
        };
    }

    if (!entity.EntityData?.SPGroupID) {
        return undefined;
    }

    return {
        id: entity.EntityData.SPGroupID,
        fullName: entity.DisplayText,
        login: entity.EntityData?.AccountName || '',
        description: entity.Description
    };
}
